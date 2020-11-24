using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GetOutlook
{
    public class GetMessages
    {
        public static List<MailMessage> Messages(MAPIFolder inboxFolder, int count, string filter, bool onlyUnread, bool markAsRead, bool getAttachments, bool TimeOrder, CancellationToken cancellationToken) 
        {
			Items items = null;
			List<MailMessage> list = new List<MailMessage>();
			try
			{
				if (inboxFolder == null)
				{
					throw new ArgumentException("MissingFolder");
				}
				bool flag = !string.IsNullOrEmpty(filter);
				bool flag2 = true;
				dynamic obj = null;
				items = inboxFolder.Items;
				items.Sort("[ReceivedTime]", TimeOrder);
				if (flag && onlyUnread)
				{
					obj = ((!filter.StartsWith("@SQL=", StringComparison.OrdinalIgnoreCase)) ? items.Find($"({filter}) and [UnRead] = True") : items.Find($"@SQL=(({filter.Substring(5)}) AND (\"urn:schemas:httpmail:read\" = 0))"));
				}
				else if (onlyUnread)
				{
					obj = items.Find($"[UnRead] = {onlyUnread}");
				}
				else if (flag)
				{
					obj = items.Find($"{filter}");
				}
				else
				{
					obj = items.GetFirst();
					flag2 = false;
				}
				int num = 0;
				while (true)
				{
					if (!((obj != null) ? true : false))
					{
						return list;
					}
					if (num == count)
					{
						return list;
					}
					if (cancellationToken.IsCancellationRequested)
					{
						break;
					}
					if (obj is MailItem)
					{
						if (markAsRead)
						{
							obj.UnRead = false;
						}
						list.Add(CreateMailMessageFromOutlookMailItem(obj, getAttachments, null));
						num++;
					}
					obj = (flag2 ? items.FindNext() : items.GetNext());
				}
				return list;
			}
			finally
			{
				if (inboxFolder != null)
				{
					Marshal.ReleaseComObject(inboxFolder);
				}
			}
		}

        private static MailMessage CreateMailMessageFromOutlookMailItem(MailItem mailItem , bool saveattachments = false, string folderpath = null)
        {
			MailMessage mailMessage = new MailMessage();
			try
			{
				try
				{
					mailMessage.Headers.Add("Uid", mailItem.EntryID);
					mailMessage.Headers["Date"] = mailItem.SentOn.ToString();
					mailMessage.Headers["DateCreated"] = mailItem.CreationTime.ToString();
					mailMessage.Headers["DateRecieved"] = mailItem.ReceivedTime.ToString();
					mailMessage.Headers["HtmlBody"] = mailItem.HTMLBody;
					mailMessage.Headers["PlainText"] = mailItem.Body;
					mailMessage.Headers["Size"] = mailItem.Size.ToString();
				}
				catch (System.Exception ex)
				{
					Trace.TraceWarning(ex.ToString());
				}
				mailMessage.Subject = mailItem.Subject;
				mailMessage.Body = mailItem.Body;
				try
				{
					string fromAddress = GetFromAddress(mailItem);
					string senderName = mailItem.SenderName;
					if (!string.IsNullOrEmpty(senderName))
					{
						mailMessage.From = new MailAddress(fromAddress, senderName);
						mailMessage.Sender = new MailAddress(fromAddress, senderName);
					}
					else
					{
						mailMessage.From = new MailAddress(fromAddress);
						mailMessage.Sender = new MailAddress(fromAddress);
					}
				}
				catch (System.Exception ex2)
				{
					Trace.TraceWarning(ex2.ToString());
				}
				mailMessage.Priority = MailPriority.Normal;
				if (mailItem.Importance == OlImportance.olImportanceHigh)
				{
					mailMessage.Priority = MailPriority.High;
				}
				if (mailItem.Importance == OlImportance.olImportanceLow)
				{
					mailMessage.Priority = MailPriority.Low;
				}
				try
				{
					if (!string.IsNullOrEmpty(mailItem.To))
					{
						foreach (MailAddress item in GetMailAddressCollection(mailItem, OlMailRecipientType.olTo))
						{
							mailMessage.To.Add(item);
						}
					}
				}
				catch (System.Exception ex3)
				{
					Trace.TraceWarning(ex3.ToString());
				}
				try
				{
					if (!string.IsNullOrEmpty(mailItem.BCC))
					{
						foreach (MailAddress item2 in GetMailAddressCollection(mailItem, OlMailRecipientType.olBCC))
						{
							mailMessage.Bcc.Add(item2);
						}
					}
				}
				catch (System.Exception ex4)
				{
					Trace.TraceWarning(ex4.ToString());
				}
				try
				{
					if (!string.IsNullOrEmpty(mailItem.CC))
					{
						foreach (MailAddress item3 in GetMailAddressCollection(mailItem, OlMailRecipientType.olCC))
						{
							mailMessage.CC.Add(item3);
						}
					}
				}
				catch (System.Exception ex5)
				{
					Trace.TraceWarning(ex5.ToString());
				}
				try
				{
					if (saveattachments && mailItem.Attachments != null)
					{
						foreach (Microsoft.Office.Interop.Outlook.Attachment attachment in mailItem.Attachments)
						{
							string text = null;
							try
							{
								text = attachment.FileName;
							}
							catch
							{
								Trace.TraceWarning("AttachmentFilenameError");
								continue;
							}
							string path = Path.Combine(folderpath ?? Path.GetTempPath(), text);
							attachment.SaveAsFile(path);
							try
							{
								mailMessage.Attachments.Add(new System.Net.Mail.Attachment(new MemoryStream(File.ReadAllBytes(path)), attachment.FileName));
							}
							finally
							{
								if (folderpath == null)
								{
									File.Delete(path);
								}
							}
						}
					}
				}
				catch (System.Exception ex6)
				{
					Trace.TraceWarning(ex6.ToString());
				}
				if (!saveattachments)
				{
					return mailMessage;
				}
				if (string.IsNullOrEmpty(mailMessage.Body))
				{
					return mailMessage;
				}
				mailMessage.AlternateViews.Add(new AlternateView(new MemoryStream(Encoding.UTF8.GetBytes(mailItem.HTMLBody)), "text/html"));
				return mailMessage;
			}
			catch (System.Exception ex7)
			{
				Trace.TraceWarning(ex7.ToString());
				return mailMessage;
			}
		}

		private static string GetFromAddress(MailItem mailItem)
		{
			Application application = null;
			try
			{
				if (!(mailItem.SenderEmailType == "EX"))
				{
					return mailItem.SenderEmailAddress;
				}
				application = GetFolder.InitOutlook();
				Recipient recipient = application.GetNamespace("MAPI").CreateRecipient(mailItem.SenderEmailAddress);
				AddressEntry addressEntry = recipient.AddressEntry;
				if (addressEntry != null)
				{
					if (addressEntry.AddressEntryUserType != 0 && addressEntry.AddressEntryUserType != OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
					{
						return recipient.Address;
					}
					ExchangeUser exchangeUser = addressEntry.GetExchangeUser();
					if (exchangeUser != null)
					{
						return exchangeUser.PrimarySmtpAddress;
					}
				}
			}
			catch (System.Exception ex)
			{
				Trace.TraceWarning(ex.ToString());
			}
			finally
			{
				if (application != null)
				{
					Marshal.ReleaseComObject(application);
				}
			}
			return string.Empty;
		}

		public static MailAddressCollection GetMailAddressCollection(MailItem mailItem, OlMailRecipientType recipentType)
		{
			MailAddressCollection mailAddressCollection = new MailAddressCollection();
			Recipients recipients = mailItem.Recipients;
			try
			{
				foreach (Recipient item in recipients)
				{
					if (item.Type == (int)recipentType)
					{
						try
						{
							string text = item.Address;
							if (!text.Contains("@"))
							{
								try
								{
									text = ((dynamic)item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")).ToString();
								}
								catch (System.Exception ex)
								{
									Trace.TraceWarning(ex.ToString());
								}
							}
							mailAddressCollection.Add(new MailAddress(text, item.Name));
						}
						catch (System.Exception ex2)
						{
							Trace.TraceWarning(ex2.ToString());
						}
					}
				}
				return mailAddressCollection;
			}
			finally
			{
				if (recipients != null)
				{
					Marshal.ReleaseComObject(recipients);
				}
			}
		}
	}
}
