class CheckMailer:

        def __init__(self, filename="LOG1.txt", mailbox="Mailbox - Another User Mailbox", folderindex=3):
            self.f = FileWriter(filename)
            self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").Folders(mailbox)
            self.inbox = self.outlook.Folders(folderindex)


        def check(self):                
        #===============================================================================
        # for i in xrange(1,100):                           #Uncomment this section if index 3 does not work for you
        #     try:
        #         self.inbox = self.outlook.Folders(i)     # "6" refers to the index of inbox for Default User Mailbox
        #         print "%i %s" % (i,self.inbox)            # "3" refers to the index of inbox for Another user's mailbox
        #     except:
        #         print "%i does not work"%i
        #===============================================================================

                self.f.pl(time.strftime("%H:%M:%S"))
                tot = 0                
                messages = self.inbox.Items
                message = messages.GetFirst()
                while message:
                    self.f.pl (message.Subject)
                    message = messages.GetNext()
                    tot += 1
                self.f.pl("Total Messages found: %i" % tot)
                self.f.pl("-" * 80)
                self.f.flush()

if __name__ == "__main__":
    mail = CheckMailer()
    for i in xrange(320):  # this is 10.6 hours approximately
            mail.check()