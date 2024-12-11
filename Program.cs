using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

/**change to desired directories for input/output**/
/**May create an executable for Windows,Mac and Linux**/
var dirPath = @"C:\Users\user\Desktop\Sent";
var outFileName = @"C:\Users\user\Desktop\outputFile.pst";

if (File.Exists(outFileName))
{
    File.Delete(outFileName);
}

using (var personalStorage = PersonalStorage.Create(outFileName, FileFormatVersion.Unicode))
{
    var inboxFolder = personalStorage.RootFolder.AddSubFolder("Inbox");

    foreach (var f in Directory.GetFiles(dirPath, "*.eml"))
    {
        using (var message = MailMessage.Load(f))
        {
            inboxFolder.AddMessage(MapiMessage.FromMailMessage(message, MapiConversionOptions.UnicodeFormat));
            Console.WriteLine($"Added File: {Path.GetFileName(f)}");
        }
    }
}

Console.WriteLine("Conversion Done!");