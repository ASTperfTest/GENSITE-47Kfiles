using System;
using System.Collections;
using NUnit.Framework;
using PlantLog.Core.Domain;
using PlantLog.Core.Persistence;

namespace PlantLog.Core.Test
{
    [TestFixture]
    public class TestOwnerDao
    {
        private IOwnerDao ownerDao = null;
        private Owner owner = null;

        [TestFixtureSetUp]
        public void TestCaseInit()
        {
            ownerDao = Utility.ApplicationContext["AdoOwnerDao"] as IOwnerDao;
            owner = new Owner();
        }

        [TestFixtureTearDown]
        public void TestCaseTearDown()
        {
        }

        [Test]
        public void Test_001_Create()
        {
            owner.OwnerId = Utility.GetGuid();
            owner.DisplayName = "Max";
            owner.Topic = "Test for NUnit";
            owner.CreatorId = "Max";
            owner.CreateDateTime = DateTime.Now;
            owner.ModifierId = "Max";
            owner.ModifyDateTime = DateTime.Now;

            Owner temp = ownerDao.Create(owner);

            Assert.AreEqual(owner.OwnerId, temp.OwnerId);
            Assert.AreEqual(owner.DisplayName, temp.DisplayName);
            Assert.AreEqual(owner.Topic, temp.Topic);
        }

        [Test]
        public void Test_002_Get()
        {
            Owner temp = ownerDao.Get(owner.OwnerId);

            Assert.AreEqual(owner.OwnerId, temp.OwnerId);
        }

        [Test]
        public void Test_003_GetAll()
        {
            IList list = ownerDao.GetAll();

            Assert.AreEqual(1, list.Count);

            Owner temp = list[0] as Owner;

            Assert.AreEqual(owner.OwnerId, temp.OwnerId);
        }

        [Test]
        public void Test_004_Update()
        {
            owner.Description = "new description";
            owner.ModifierId = "Max 2";
            owner.ModifyDateTime = DateTime.Now;

            Owner temp = ownerDao.Update(owner);

            Assert.AreEqual(owner.Description, temp.Description);
            Assert.AreEqual(owner.ModifierId, temp.ModifierId);
        }

        [Test]
        public void Test_005_Delete()
        {
            ownerDao.Delete(owner.OwnerId);
        }
    }
}
