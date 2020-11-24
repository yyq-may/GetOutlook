using Microsoft.Office.Interop.Outlook;
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
    public class GetOutlookMessage : GetMailActivity
    {
        [Category("Input")]
        [DisplayName("KnownFolderDisplayName")]
        [Browsable(false)]
        [DefaultValue(KnownFolders.None)]
        public KnownFolders KnownFolder { get; set; }


        [Category("输入")]
        [DisplayName("邮件文件夹")]
        public InArgument<string> MailFolder { get; set; }

        [Category("输入")]
        [DisplayName("账户")]
        public InArgument<string> Account { get; set; }

        [Category("选项")]
        [DisplayName("筛选")]
        public InArgument<string> Filter { get; set; }

        [Category("选项")]
        [DisplayName("仅限未读消息")]
        public bool OnlyUnreadMessages { get; set; }

        [Category("选项")]
        [DisplayName("标记为已读")]
        public bool MarkAsRead { get; set; }


        [Category("选项")]
        [DisplayName("时间正序")]
        public bool TimeOrder { get; set; }
        
        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (KnownFolder != 0 && (MailFolder == null || MailFolder.Expression == null))
            {
                MailFolder = new InArgument<string>(KnownFolder.ToString());
                KnownFolder = KnownFolders.None;
            }
            base.CacheMetadata(metadata);
        }

        public GetOutlookMessage()
        {
            MailFolder = "Inbox";
            OnlyUnreadMessages = true;
            TimeOrder = true;
        }


        protected override Task<List<MailMessage>> GetMessage(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            string account = Account.Get(context);
            string filter = Filter.Get(context);
            string folderpath = MailFolder.Get(context);
            int top = Top.Get(context);
            if (KnownFolder != 0)
            {
                KnownFolders result = KnownFolders.None;
                if (!Enum.TryParse(folderpath, ignoreCase: true, out result))
                {
                    folderpath = KnownFolder.ToString() + "\\" + folderpath;
                }
            }
            MAPIFolder mAPIFolder = GetFolder.GetFolders(folderpath, account);
            List<MailMessage> mailMessages = GetMessages.Messages(mAPIFolder, top, filter, OnlyUnreadMessages, MarkAsRead, true, TimeOrder,cancellationToken);            
            return StartNew(() => mailMessages);
        }
        public static Task<TResult> StartNew<TResult>(Func<TResult> func)
        {
            return Task.Factory.StartNew(func, CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default);
        }
        protected override void TaskHandler(AsyncCallback callback, Task<List<MailMessage>> task, TaskCompletionSource<List<MailMessage>> tcs, CancellationToken token, int timeout)
        {
            TaskHandlers(callback, task, tcs, token, timeout);
        }

        public static void TaskHandlers(AsyncCallback callback, Task<List<MailMessage>> task, TaskCompletionSource<List<MailMessage>> tcs, CancellationToken token, int timeout)
        {
            timeout = ((timeout <= 0) ? 30000 : timeout);
            Task.Run(delegate
            {
                try
                {
                    if (!task.Wait(timeout, token))
                    {
                        tcs.TrySetException(new TimeoutException());
                    }
                    else if (token.IsCancellationRequested || task.IsCanceled)
                    {
                        tcs.TrySetCanceled();
                    }
                    else if (task.IsFaulted)
                    {
                        tcs.TrySetException(task.Exception.InnerExceptions);
                    }
                    else
                    {
                        tcs.TrySetResult(task.Result);
                    }
                }
                catch (System.Exception ex)
                {
                    tcs.TrySetException(ex.InnerException);
                }
                callback?.Invoke(tcs.Task);
            });
        }
    }
}
