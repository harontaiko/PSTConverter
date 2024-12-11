using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

/**change to desired directories for input/output**/
/**May create an executable for Windows,Mac and Linux**/
var dirPath = @"C:\Users\user\Desktop\Notification2";
var outFileName = @"C:\Users\user\Desktop\Notification2.pst";

if (File.Exists(outFileName))
{
    File.Delete(outFileName);
}

using (var personalStorage = PersonalStorage.Create(outFileName, FileFormatVersion.Unicode))
{
    var inboxFolder = personalStorage.RootFolder.AddSubFolder("Notification2");
    int index = 1;

    foreach (var f in Directory.GetFiles(dirPath, "*.eml"))
    {
        try
        {
            using (var message = MailMessage.Load(f))
            {
                inboxFolder.AddMessage(MapiMessage.FromMailMessage(message, MapiConversionOptions.UnicodeFormat));
                Console.WriteLine($"Added File {index}: {Path.GetFileName(f)}");
            }
        }
        catch (System.UriFormatException ex)
        {
            Console.WriteLine($"Skipping file {index}: {Path.GetFileName(f)} due to error: {ex.Message}");
        }
        finally
        {
            index++; 
        }
    }
}

Console.WriteLine("Conversion Done!");