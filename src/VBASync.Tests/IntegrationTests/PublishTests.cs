using NUnit.Framework;
using System.Linq;
using VBASync.Model;
using VBASync.Tests.Mocks;

namespace VBASync.Tests.IntegrationTests
{
    [TestFixture]
    public class PublishTests
    {
        [Test]
        public void PublishToXlsmIsRepeatable()
        {
            var fso = new FakeSystemOperations();
            fso.FileWriteAllBytes("Book1.xlsm", Files.PublishToXlsmIsRepeatableFiles.Book1);
            fso.FileWriteAllBytes("repo/Module1.bas", Files.PublishToXlsmIsRepeatableFiles.Module1);
            fso.FileWriteAllBytes("repo/Project.ini", Files.PublishToXlsmIsRepeatableFiles.Project);
            fso.FileWriteAllBytes("repo/Sheet1.cls", Files.PublishToXlsmIsRepeatableFiles.Sheet1);
            fso.FileWriteAllBytes("repo/Sheet2.cls", Files.PublishToXlsmIsRepeatableFiles.Sheet2);
            fso.FileWriteAllBytes("repo/ThisWorkbook.cls", Files.PublishToXlsmIsRepeatableFiles.ThisWorkbook);

            var session = new QuickSession
            {
                Action = ActionType.Publish,
                AutoRun = true,
                FilePath = "Book1.xlsm",
                FolderPath = "repo/"
            };
            var settings = new QuickSessionSettings
            {
                AddNewDocumentsToFile = false,
                AfterExtractHook = null,
                BeforePublishHook = null,
                DeleteDocumentsFromFile = false,
                IgnoreEmpty = false
            };

            using (var actor = new ActiveSession(fso, session, settings))
            {
                actor.Apply(actor.GetPatches().ToList());
            }
            var fileContentsAfterPublish1 = fso.FileReadAllBytes("Book1.xlsm");

            fso.FileWriteAllBytes("Book1.xlsm", Files.PublishToXlsmIsRepeatableFiles.Book1);

            using (var actor = new ActiveSession(fso, session, settings))
            {
                actor.Apply(actor.GetPatches().ToList());
            }
            var fileContentsAfterPublish2 = fso.FileReadAllBytes("Book1.xlsm");

            Assert.That(fileContentsAfterPublish1, Is.EqualTo(fileContentsAfterPublish2));
        }
    }
}
