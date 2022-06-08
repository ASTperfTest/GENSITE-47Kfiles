using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PlantLog.Core.Persistence;
using PlantLog.Core.Domain;
using NUnit.Framework;
using System.Collections;

namespace PlantLog.Core.Test
{
    [TestFixture]
    public class TestVoteRecordDao
    {
        private IVoteRecordDao voteRecordDao = null;
        private VoteRecord voteRecord = null;

        [TestFixtureSetUp]
        public void TestCaseInit()
        {
            voteRecordDao = Utility.ApplicationContext["AdoVoteRecordDao"] as IVoteRecordDao;
            voteRecord = new VoteRecord();
        }

        [TestFixtureTearDown]
        public void TestCaseTearDown()
        {
        }

        [Test]
        public void Test_001_Create()
        {
            voteRecord.IP = "172.16.5.0";
            voteRecord.VoteDate = DateTime.Now;
            voteRecord.VoterId = "Max";
            voteRecord.UserId = "Woods";

            voteRecordDao.Create(voteRecord);
        }

        [Test]
        public void Test_002_GetByUser()
        {
            IList temp = voteRecordDao.GetByUser(voteRecord.UserId);

            Assert.IsTrue(temp.Count > 0);
        }

        [Test]
        public void Test_003_GetAll()
        {
            IList temp = voteRecordDao.GetAll();

            Assert.IsTrue(temp.Count > 0);
        }

        [Test]
        public void Test_004_Update()
        {
            voteRecord.VoteDate = DateTime.Now;
            voteRecord.UserId = "Woods";

            voteRecordDao.Update(voteRecord);
        }

        [Test]
        public void Test_005_GetVoteRecordByVoter()
        {
            IList temp = voteRecordDao.GetVoteRecordByVoter(voteRecord.VoterId);

            Assert.AreEqual(1, temp.Count);
        }
    }
}
