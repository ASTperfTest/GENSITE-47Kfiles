using System;
using System.Collections;
using NUnit.Framework;
using Gardening.Core.Domain;
using Gardening.Core.Persistence;

namespace Gardening.Core.Test
{
    [TestFixture]
    public class TestEntryDao
    {
        private IEntryDao entryDao = null;
        private Entry entry = null;
        private string topicId = Utility.GetGuid();
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
            entry.TopicId = topicId;
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
            Assert.AreEqual(entry.TopicId, temp.TopicId);
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
            Assert.AreEqual(entry.TopicId, temp.TopicId);
            Assert.AreEqual(entry.Title, temp.Title);
            Assert.IsTrue(temp.IsPublic);
        }

        [Test]
        public void Test_003_GetByTopic()
        {
            IList list = entryDao.GetByTopic(entry.TopicId);

            Assert.AreEqual(1, list.Count);

            Entry temp = list[0] as Entry;

            Assert.AreEqual(entry.Date, temp.Date);
            Assert.AreEqual(entry.Description, temp.Description);
            Assert.AreEqual(entry.EntryId, temp.EntryId);
            Assert.AreEqual(entry.TopicId, temp.TopicId);
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
            Assert.AreEqual(entry.TopicId, temp.TopicId);
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
            Assert.AreEqual(entry.TopicId, temp.TopicId);
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
            entryDao.DeleteByTopic(entry.TopicId);
            IList list = entryDao.GetAll();
            Assert.AreEqual(0, list.Count);
        }
    }
}
