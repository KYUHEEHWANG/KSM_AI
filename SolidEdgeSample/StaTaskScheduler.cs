using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace KSM_SolidEdge
{
    /// <summary>
    /// 단일 STA 스레드에서 Task를 실행합니다. COM(예: Solid Edge) 자동화는 동일 아파트먼트에서만 호출해야 합니다.
    /// </summary>
    public sealed class StaTaskScheduler : TaskScheduler, IDisposable
    {
        private static readonly Lazy<StaTaskScheduler> LazyInstance =
            new Lazy<StaTaskScheduler>(() => new StaTaskScheduler());

        public static StaTaskScheduler Instance => LazyInstance.Value;

        private readonly BlockingCollection<Task> _tasks = new BlockingCollection<Task>();
        private readonly Thread _thread;
        private bool _disposed;

        private StaTaskScheduler()
        {
            _thread = new Thread(StaThreadLoop)
            {
                Name = "SolidEdge STA",
                IsBackground = true
            };
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        private void StaThreadLoop()
        {
            foreach (var task in _tasks.GetConsumingEnumerable())
            {
                TryExecuteTask(task);
            }
        }

        protected override IEnumerable<Task> GetScheduledTasks()
        {
            return _tasks.ToArray();
        }

        protected override void QueueTask(Task task)
        {
            _tasks.Add(task);
        }

        protected override bool TryExecuteTaskInline(Task task, bool taskWasPreviouslyQueued)
        {
            return false;
        }

        public override int MaximumConcurrencyLevel => 1;

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _tasks.CompleteAdding();
            try
            {
                if (_thread.IsAlive)
                    _thread.Join(TimeSpan.FromSeconds(10));
            }
            catch { }
        }
    }
}
