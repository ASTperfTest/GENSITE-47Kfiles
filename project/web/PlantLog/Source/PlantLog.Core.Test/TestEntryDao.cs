using System;
using System.Collections;
using NUnit.Framework;
using PlantLog.Core.Domain;
using PlantLog.Core.Persistence;

namespace PlantLog.Core.Test
{
    [TestFixture]
    public class TestEntryDao
    {
        private IEntryDao entryDao = null;
        private Entry entry = null;

        [TestFixtureSetUp]
        public void TestCaseInit()
        {
            entryDao = Utility.ApplicationContext["AdoEntryDao"] as IEntryDao;
            entry = new Entry();
        }

        [TestFixtureTearDown]
        public void TestCaseTearDown()
        {
        }

        [Test]
        public void Test_001_Create()
        {
            entry.Date = DateTime.Today;
            entry.Description = "new entry description";
            entry.EntryId = Utility.GetGuid();
            entry.OwnerId = Utility.GetGuid();
            entry.Title = "test for new";
            entry.IsPublic = true;
            entry.CreatorId = "Max";
            entry.CreateDateTime = DateTime.Now;
            entry.ModifierId = "Max";
            entry.ModifyDateTime = DateTime.Now;

            Entry temp = entryDao.Create(entry);

            Assert.AreEqual(entry.Date, temp.Date);
            Assert.AreEqual(entry.Description, temp.Description);
            Assert.AreEqual(entry.EntryId, temp.EntryId);
            Assert.AreEqual(entry.OwnerId, temp.OwnerId);
            Assert.AreEqual(entry.Title, temp.Title);
            Assert.IsTrue(temp.IsPublic);
        }

        [Test]
        public void Test_002_Get()
        {
            Entry temp = entryDao.Get(entry.EntryId);

            Assert.AreEqual(entry.Date, temp.Date);
            Assert.AreEqual(entry.Description, temp.Description);
            Assert.AreEqual(entry.EntryId, temp.EntryId);
            Assert.AreEqual(entry.OwnerId, temp.OwnerId);
            Assert.AreEqual(entry.Title, temp.Title);
            Assert.IsTrue(temp.IsPublic);
        }

        [Test]
        public void Test_003_GetByOwner()
        {
            IList list = entryDao.GetByOwner(entry.OwnerId);

            Assert.AreEqual(1, list.Count);

            Entry temp = list[0] as Entry;

            Assert.AreEqual(entry.Date, temp.Date);
            Assert.AreEqual(entry.Description, temp.Description);
            Assert.AreEqual(entry.EntryId, temp.EntryId);
            Assert.AreEqual(entry.OwnerId, temp.OwnerId);
            Assert.AreEqual(entry.Title, temp.Title);
            Assert.IsTrue(temp.IsPublic);
        }

        [Test]
        public void Test_004_GetAll()
        {
            IList list = entryDao.GetAll();
            Assert.AreEqual(1, list.Count);
            Entry temp = list[0] as Entry;

            Assert.AreEqual(entry.Date, temp.Date);
            Assert.AreEqual(entry.Description, temp.Description);
            Assert.AreEqual(entry.EntryId, temp.EntryId);
            Assert.AreEqual(entry.OwnerId, temp.OwnerId);
            Assert.AreEqual(entry.Title, temp.Title);
            Assert.IsTrue(temp.IsPublic);
        }

        [Test]
        public void Test_005_Update()
        {
            entry.Title = "Update Entry Title";
            entry.Description = "Update new description";
            entry.IsPublic = false;
            entry.ModifierId = "Max";
            entry.ModifyDateTime = DateTime.Now;

            Entry temp = entryDao.Update(entry);

            Assert.AreEqual(entry.Date, temp.Date);
            Assert.AreEqual(entry.Description, temp.Description);
            Assert.AreEqual(entry.EntryId, temp.EntryId);
            Assert.AreEqual(entry.OwnerId, temp.OwnerId);
            Assert.AreEqual(entry.Title, temp.Title);
            Assert.AreEqual(entry.ModifierId, temp.ModifierId);
            Assert.IsFalse(temp.IsPublic);
        }

        [Test]
        public void Test_006_Delete()
        {
            entryDao.Delete(entry.EntryId);
        }

        [Test]
        public void Test_007_DeleteByOwner()
        {
            entryDao.DeleteByOwner(entry.OwnerId);
            IList list = entryDao.GetAll();
            Assert.AreEqual(0, list.Count);
        }
    }
}
