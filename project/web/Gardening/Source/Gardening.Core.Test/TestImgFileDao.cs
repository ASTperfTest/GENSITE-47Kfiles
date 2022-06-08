using System.Collections;
using System.IO;
using NUnit.Framework;
using Gardening.Core.Domain;
using Gardening.Core.Persistence;
using System;

namespace Gardening.Core.Test
{
    [TestFixture]
    public class TestImgFileDao
    {
        private IImgFileDao sourcefileDao = null;
        private IEntryDao entryDao = null;
        private Entry entry = null;
        private ImgFile file = null;
        private FileInfo testFile = null;

        [TestFixtureSetUp]
        public void TestCaseInit()
        {
            entryDao = Utility.ApplicationContext["AdoEntryDao"] as IEntryDao;
            sourcefileDao = Utility.ApplicationContext["AdoImgFileDao"] as IImgFileDao;
            file = new ImgFile();
            testFile = new FileInfo("E:\\Download\\test.txt");

            entry = new Entry();
            entry.Date = DateTime.Today;
            entry.Description = "new entry description";
            entry.EntryId = Utility.GetGuid();
            entry.TopicId = Utility.GetGuid();
            entry.Title = "test for new";
            entry.CreatorId = "Max";
            entry.CreateDateTime = DateTime.Now;
            entry.ModifierId = "Max";
            entry.ModifyDateTime = DateTime.Now;

            entryDao.Create(entry);
        }

        [TestFixtureTearDown]
        public void TestCaseTearDown()
        {
            entryDao.Delete(entry.EntryId);
        }

        [Test]
        public void Test_001_Create()
        {
            file.ParentId = entry.EntryId;
            file.Name = testFile.Name;
            file.FileId = Utility.GetGuid();
            file.Uri = testFile.DirectoryName;

            ImgFile temp = sourcefileDao.Create(file);

            Assert.AreEqual(file.ParentId, temp.ParentId);
            Assert.AreEqual(file.Name, temp.Name);
            Assert.AreEqual(file.FileId, temp.FileId);
            Assert.AreEqual(file.Uri, temp.Uri);
        }

        [Test]
        public void Test_002_Get()
        {
            ImgFile temp = sourcefileDao.Get(file.FileId);

            Assert.AreEqual(file.ParentId, temp.ParentId);
            Assert.AreEqual(file.Name, temp.Name);
            Assert.AreEqual(file.FileId, temp.FileId);
            Assert.AreEqual(file.Uri, temp.Uri);
        }

        [Test]
        public void Test_003_GetByEntry()
        {
            IList list = sourcefileDao.GetByParent(file.ParentId);

            Assert.AreEqual(1, list.Count);

            ImgFile temp = list[0] as ImgFile;

            Assert.AreEqual(file.ParentId, temp.ParentId);
            Assert.AreEqual(file.Name, temp.Name);
            Assert.AreEqual(file.FileId, temp.FileId);
            Assert.AreEqual(file.Uri, temp.Uri);
        }

        [Test]
        public void Test_004_Update()
        {
            file.Name = "Update File Name";

            ImgFile temp = sourcefileDao.Update(file);

            Assert.AreEqual(file.ParentId, temp.ParentId);
            Assert.AreEqual(file.Name, temp.Name);
            Assert.AreEqual(file.FileId, temp.FileId);
            Assert.AreEqual(file.Uri, temp.Uri);
        }

        [Test]
        public void Test_005_Delete()
        {
            sourcefileDao.Delete(file.FileId);
        }

        [Test]
        public void Test_006_DeleteByEntry()
        {
            sourcefileDao.DeleteByParent(file.ParentId);
        }
    }
}
