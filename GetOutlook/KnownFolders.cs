using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetOutlook
{
    public enum KnownFolders
    {
		None = 0,
	    Inbox = 6,
	    SentMail = 5,
	    Outbox = 4,
	    Junk = 23,
	    Drafts = 0x10,
	    DeletedItems = 3
    }
}
