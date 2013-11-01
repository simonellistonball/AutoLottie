using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Diagnostics;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Threading = System.Threading;


namespace OutlookAddIn1
{
    public partial class AutoLottieAddIn
    {
        // Time from original Sent Time until the reminder should be sent
        public static String LottieDurationProperty = "AutoLottieDuration";

        // Set this property on a message so that receiving the message does NOT cancel the thread.
        // Used for the start of the thread and Lottie Reminders.
        // Arguably, it should be per-person reminders and only cancelled per-person, but shrug.
        public static String LottieNoCancelProperty = "AutoLottieNoCancel";

        // Approximate Send Time.  This will go away if we use the real message from sent-mail
        public static String ApproxSentTime = "SentApproximatelyAt";

        Outlook.MAPIFolder lottieFolder;

        static Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;

        // if we don't declare an Items collection to add the event handler to, it gets garbage collected away.
        Outlook.Items inboxItems;

        Dictionary<String, Threading.Timer> messagesAndTimers;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           // In case you need to attach a debugger.
            // Commented out for Demo - normally this would be handled by DebugBox, but 
 //           DebugBox("halt!");

            // Some Initializations
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            // lame that we have to do this. lame lame lame. Just using inbox.Items and the event handler gets garbage collected.
            inboxItems = inbox.Items;
            // Add a callback to check every time a message is received whether it is one we've been waiting for
            inboxItems.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(CheckLottie);

            // Add a callback on send to put the message in the right place
            Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(FireBeforeSent);

            // Find the Folder of Items for follow-up
            lottieFolder = FindOrCreateLottieFolder();

            messagesAndTimers = new Dictionary<string,Threading.Timer>();
            // Go through the messages in the folder and put some timers in for their follow-up
            foreach (Outlook.MailItem child in lottieFolder.Items)
            {
                AddTimerForMessage(child);
            }
        }

        private void AddTimerForMessage(Outlook.MailItem child)
        {
            try
            {
                TimeSpan expiryTimeFromSent;
                TimeSpan.TryParse(child.UserProperties[LottieDurationProperty].Value.ToString(), out expiryTimeFromSent);
                DateTime sentTime;
                DateTime.TryParse(child.UserProperties[ApproxSentTime].Value.ToString(), out sentTime);


                TimeSpan timeToNag = sentTime.Subtract(DateTime.Now) + expiryTimeFromSent;
                // If timer has expired or is set to go within the next minute, do it in a minute. - see below for the note on nagging while Outlook
                // is starting up loading e-mail which >might< contain a response!
                if (timeToNag.CompareTo(new TimeSpan(0, 1, 0)) < 1)
                {
                    //Used to do the Nag immediately, but  if someone is firing up Outlook and this goes before the email has been loaded....
                    // So now we're just going to nag in 1 minute
                    timeToNag = new TimeSpan(0, 1, 0);
                }

                // Add a Timer to fire off the nag when appropriate
                Threading.TimerCallback nagCallBack = DoNag;
                Threading.Timer t = new Threading.Timer(nagCallBack, child, timeToNag, new TimeSpan(0));

                // Keep a table of some id and the timers so that we can cancel it if the person does reply
                messagesAndTimers.Add(child.ConversationID, t);

                DebugBox("Scheduling nag for " + timeToNag.ToString());

            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not schedule followup for message " + child.Subject + ".");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        public void DoNag(Object item)
        {
            Outlook.MailItem msg = (Outlook.MailItem)item;
            // Blah.  Creating a new mail from scratch is a problem, because if the user then replies to the AutoLottie, it doesn't 
            // hook up.  Creating a new mail using 'reply' is impossible because the one in the Lottie folder is still only a draft.
            // Going to try looking up the e-mail in regular sent-mail... If this works, then longer term, the approach should be 
            // changed so that the real e-mail gets stored in the Lottie folder, introducing the requirement that mail goes to sent-mail.
            // Also, we can trigger the Lottie actions with a watch on sent-mail, which has a lot of other benefits.
            Outlook.MAPIFolder sentmail = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
            // TODO FALLS OVER WITH PUNCTUATION (like ') in the SUBJECT LINE
            String sFilter = "[Subject] = '"+ msg.Subject +"'";
            Outlook.Items sublist = sentmail.Items.Restrict(sFilter);
            DateTime latestSentTime = new DateTime(0);
            Outlook.MailItem bestFitMail = null;
            Outlook.MailItem messageIt = (Outlook.MailItem)sublist.GetFirst();
            while (messageIt != null)
            {

                if ((((Outlook.MailItem)messageIt).ConversationID == msg.ConversationID) && (((Outlook.MailItem)messageIt).SentOn.CompareTo(latestSentTime) == 1))
                {
                    bestFitMail = (Outlook.MailItem)messageIt;
                    latestSentTime = ((Outlook.MailItem)messageIt).SentOn;
                }
                messageIt = (Outlook.MailItem)sublist.GetNext();
            }

            if (bestFitMail != null)
            {
                Outlook.MailItem eMail = bestFitMail.ReplyAll();
                //(Outlook.MailItem)this.Application.CreateItem(Outlook.OlItemType.olMailItem);
                // if you change the subject, Conversation ID gets changed :-(
             //   eMail.Subject = "[AutoLottie] " + bestFitMail.Subject;
                //          eMail.To = msg.To;
                eMail.CC = bestFitMail.SenderEmailAddress;
                eMail.Body = "As far as my Outlook client can tell, you have not yet responded " +
                        "to the message below.  Could you do so now, please?\n\n " +
                        "Thanks! \n\n" +
                        "************************************************************\n\n\n" + bestFitMail.Body;
                eMail.Importance = Outlook.OlImportance.olImportanceNormal;
               
                Outlook.UserProperty lottieStartThreadProperty = eMail.UserProperties.Add(AutoLottieAddIn.LottieNoCancelProperty, Outlook.OlUserPropertyType.olText);
                lottieStartThreadProperty.Value = "True";

                ((Outlook._MailItem)eMail).Send();
            }
            // Not sure what to do here:
            // 1) Set the AutoLottie property so that they get nagged again in the same length of time?
            // 2) Remove the AutoLottie?
            // 3) Something random - they get nagged in double the time.  I think I'll do that. :-)
            TimeSpan expiryTimeFromSent;
            TimeSpan.TryParse(msg.UserProperties[LottieDurationProperty].Value.ToString(), out expiryTimeFromSent);
            TimeSpan newExpiryTime = new TimeSpan(2 * expiryTimeFromSent.Days, 2 * expiryTimeFromSent.Hours, 2 * expiryTimeFromSent.Minutes, 0);

            UpdateTimerForMessage(msg, newExpiryTime);

        }

        private void UpdateTimerForMessage(Outlook.MailItem msg, TimeSpan when)
        {
            Threading.Timer theTimer;
            messagesAndTimers.TryGetValue(msg.ConversationID, out theTimer);

            messagesAndTimers.Remove(msg.ConversationID);
            messagesAndTimers.Add(msg.ConversationID, theTimer);
            theTimer.Change(when, when);
   
        }

        private Outlook.MAPIFolder FindOrCreateLottieFolder()
        {
            // Create new AutoLottie folder if necessary
            Outlook.MAPIFolder sentmail = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
            Outlook.Folders childFolders = sentmail.Folders;
            foreach (Outlook.Folder child in childFolders)
            {
                if (child.Name.Equals("AutoLottie"))
                {
                    return child;
                }
            }

            //create 
            return sentmail.Folders.Add("AutoLottie");

        }

        // For each message in the Lottie folder, check to see if it's part of the same conversation as the message that
        // just arrived.  If so, remove the nag.  Hooray!
        // Limitations:  Don't Lottie more than one e-mail in a conversation.
        private void CheckLottie(object stateInfo)
        {
            // TODO if this is a calendar item, obviously the cast fails.  oops.
            Outlook.MailItem msg = (Outlook.MailItem)stateInfo;
            // remove corresponding timer
            Threading.Timer myTimer;
            String check = msg.ConversationID;
            if (messagesAndTimers.TryGetValue(msg.ConversationID, out myTimer))
            {
                // right conversation. Check to see if we >shouldn't< cancel the nag.

                if (msg.UserProperties[LottieNoCancelProperty] != null)
                {
                    String inString = msg.UserProperties[LottieNoCancelProperty].Value;
                    if (inString.Equals("True"))
                    {
                        return;
                    }
                }

                // Hooray!  Don't Nag!
                myTimer.Dispose();
                DebugBox("Deleting nag for " + msg.Subject);

                // obviously, this goes wrong if you've Lottied more than one in the same conversation.
                messagesAndTimers.Remove(msg.ConversationID);

                // find the Lottie message in the sent mail->AutoLottie folder and remove it.
                foreach (System.Object curItem in lottieFolder.Items)
                {
                    try
                    {
                        Outlook.MailItem message = (Outlook.MailItem)curItem;
                        if (message.ConversationID == msg.ConversationID)
                        {
                            message.Delete();
                            break;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }


            }

        }

        // yes, unfortunately, this fires >before< the mail is sent. This means we can't just use the sent message, and it also has the unfortuante consequence
        // that the thing that goes into the AutoLottie folder is still a draft.  An alternative could be to watch the sent-mail folder. I didn't do this because 
        // not everyone uses sent-mail, but that may be a wrong call. 

        private void FireBeforeSent(object Item, ref bool Cancel)
        {
            var msg = Item as Outlook.MailItem;
            if (msg.UserProperties[LottieDurationProperty] != null)
            {
                Outlook.MailItem copiedMsg = msg.Copy();

                // One of the consequences of not having a sent time is that, well, we don't know when it was sent.  Storing an approximate time.
                Outlook.UserProperty approxSentTime = copiedMsg.UserProperties.Add(ApproxSentTime, Outlook.OlUserPropertyType.olText);
                approxSentTime.Value = DateTime.Now.ToString();
                copiedMsg.Save();
                copiedMsg.Move(lottieFolder);
                //commented out for demo
             //   DebugBox(copiedMsg.Subject + " is " + copiedMsg.UserProperties[ApproxSentTime].Value);

                // set the timer going for this in the current session
                AddTimerForMessage(copiedMsg);
            }
        }

       // [Conditional("DEBUG")]  <- why doesn't this work?  don't know.
        private void DebugBox(String s)
        {
            MessageBox.Show(s);
        }
        
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }



        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
