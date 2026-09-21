using System.Windows.Threading;

namespace Kanlook.Services;

/// <summary>
/// A single background STA thread that every Outlook COM call runs on.
///
/// Outlook automation is apartment-bound, which used to mean "run it on the WPF UI thread" - and a
/// mailbox big enough to take minutes to enumerate then froze the window for those minutes. Giving
/// COM a thread of its own keeps the UI responsive; the work is no faster, it just happens
/// somewhere else. The thread's dispatcher queue also serialises the calls, so the COM objects are
/// still only ever touched by one thread at a time and need no locking of their own.
/// </summary>
public sealed class ComWorker : IDisposable
{
    private readonly Dispatcher _dispatcher;
    private readonly Thread _thread;

    public ComWorker()
    {
        var ready = new TaskCompletionSource<Dispatcher>(TaskCreationOptions.RunContinuationsAsynchronously);

        _thread = new Thread(() =>
        {
            ready.SetResult(Dispatcher.CurrentDispatcher);
            Dispatcher.Run();
        })
        {
            Name = "Outlook COM",
            IsBackground = true,
        };

        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();

        _dispatcher = ready.Task.GetAwaiter().GetResult();
    }

    /// <summary>
    /// Queues work onto the COM thread. Background-priority work yields to everything the user is
    /// waiting on, so a long indexing pass can't get in front of opening a folder or a mail.
    /// </summary>
    public Task<T> RunAsync<T>(Func<T> work, DispatcherPriority priority = DispatcherPriority.Normal) =>
        _dispatcher.InvokeAsync(work, priority).Task;

    public Task RunAsync(Action work, DispatcherPriority priority = DispatcherPriority.Normal) =>
        _dispatcher.InvokeAsync(work, priority).Task;

    public void Dispose()
    {
        _dispatcher.InvokeShutdown();

        // Bounded: a call stuck inside Outlook must not keep the process alive on exit. The thread
        // is a background one, so the runtime tears it down either way.
        _thread.Join(TimeSpan.FromSeconds(5));
    }
}
