using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GetOutlook
{
    public abstract class GetMailActivity : AsyncCodeActivity
    {
        [Category("常见")]
        [DisplayName("超时")]
        public InArgument<int> TimeoutMS { get; set; }

        [Category("选项")]
        [DisplayName("读取的个数")]
        public InArgument<int> Top { get; set; }

        [Category("输出")]
        [DisplayName("消息")]
        public OutArgument<List<MailMessage>> Messages { get; set; }

        protected GetMailActivity()
        {
            Top = new InArgument<int>(30);
        }

        protected override IAsyncResult BeginExecute(AsyncCodeActivityContext context, AsyncCallback callback, object state)
        {
            CancellationTokenSource cancellationTokenSource2 = (CancellationTokenSource)(context.UserState = new CancellationTokenSource());
            int timeout = TimeoutMS.Get(context);
            Task<List<MailMessage>> messages = GetMessage(context, cancellationTokenSource2.Token);
            TaskCompletionSource<List<MailMessage>> taskCompletionSource = new TaskCompletionSource<List<MailMessage>>(state);
            TaskHandler(callback, messages, taskCompletionSource, cancellationTokenSource2.Token, timeout);
            return taskCompletionSource.Task;
        }
        protected abstract Task<List<MailMessage>> GetMessage(AsyncCodeActivityContext context, CancellationToken cancellationToken);

        protected override void Cancel(AsyncCodeActivityContext context)
        {
            ((CancellationTokenSource)context.UserState).Cancel();
        }
        protected override void EndExecute(AsyncCodeActivityContext context, IAsyncResult result)
        {
            Task<List<MailMessage>> task = (Task<List<MailMessage>>)result;
            try
            {
                if (task.IsFaulted)
                {
                    throw task.Exception.InnerException;
                }
                if (task.IsCanceled || context.IsCancellationRequested)
                {
                    context.MarkCanceled();
                }
                else
                {
                    Messages.Set(context, task.Result);
                }
            }
            catch (OperationCanceledException)
            {
                context.MarkCanceled();
            }
        }
        protected virtual void TaskHandler(AsyncCallback callback, Task<List<MailMessage>> task, TaskCompletionSource<List<MailMessage>> tcs, CancellationToken token, int timeout)
        {
            task.ContinueWith(delegate (Task<List<MailMessage>> t)
            {
                if (token.IsCancellationRequested || t.IsCanceled)
                {
                    tcs.TrySetCanceled();
                }
                else if (t.IsFaulted)
                {
                    tcs.TrySetException(t.Exception.InnerExceptions);
                }
                else
                {
                    tcs.TrySetResult(t.Result);
                }
                callback?.Invoke(tcs.Task);
            });
        }
    }
}
