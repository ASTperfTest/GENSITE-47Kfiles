using System;
using System.Collections;
using NUnit.Framework;
using Gardening.Core.Domain;
using Gardening.Core.Persistence;

namespace Gardening.Core.Test
{
    [TestFixture]
    public class TestTopicDao
    {
        private ITopicDao topicDao = null;
        private Topic topic = null;

        [TestFixtureSetUp]
        public void TestCaseInit()
        {
            topicDao = Utility.ApplicationContext["AdoTopicDao"] as ITopicDao;
            topic = new Topic();
        }

        [TestFixtureTearDown]
        public void TestCaseTearDown()
        {
            IList topicList = topicDao.GetAll();
            foreach (Topic topic in topicList)
            {
                topicDao.Delete(topic.TopicId);
            }
        }

        [Test]
        public void Test_001_Create()
        {
            topic.Title = "Test for NUnit";
            User owner = new User();
            owner.UserId = "Max";
            topic.Owner = owner;
            topic.Owner.DisplayName = "Max";
            topic.CreateDateTime = DateTime.Now;
            topic.ModifierId = "Max";
            topic.ModifyDateTime = DateTime.Now;
            topic.RecommendedOrder = 1;

            Topic temp = topicDao.Create(topic);

            Assert.AreEqual(topic.Owner.UserId, temp.Owner.UserId);
            Assert.AreEqual(topic.Owner.DisplayName, temp.Owner.DisplayName);
            Assert.AreEqual(topic.Title, temp.Title);
            Assert.AreEqual(topic.RecommendedOrder, temp.RecommendedOrder);
        }

        [Test]
        public void Test_002_Get()
        {
            Topic temp = topicDao.Get(topic.TopicId);

            Assert.AreEqual(topic.Owner.UserId, temp.Owner.UserId);
            Assert.AreEqual(topic.RecommendedOrder, temp.RecommendedOrder);
        }

        [Test]
        public void Test_003_GetAll()
        {
            IList list = topicDao.GetAll();

            Assert.AreEqual(1, list.Count);

            Topic temp = list[0] as Topic;

            Assert.AreEqual(topic.Owner.UserId, temp.Owner.UserId);
            Assert.AreEqual(topic.RecommendedOrder, temp.RecommendedOrder);
        }

        [Test]
        public void Test_004_Update()
        {
            topic.Description = "new description";
            topic.ModifierId = "Max 2";
            topic.ModifyDateTime = DateTime.Now;
            topic.RecommendedOrder = 2;
            Topic temp = topicDao.Update(topic);

            Assert.AreEqual(topic.Description, temp.Description);
            Assert.AreEqual(topic.ModifierId, temp.ModifierId);
            Assert.AreEqual(topic.RecommendedOrder, temp.RecommendedOrder);
        }

        [Test]
        public void Test_005_GetByRecommendedOrder()
        {
            IList il = topicDao.GetByRecommendedOrder(3);
            Assert.AreEqual(1, il.Count);
            Assert.AreEqual(((Topic)il[0]).RecommendedOrder, 2);
        }

        [Test]
        public void Test_006_Delete()
        {
            topicDao.Delete(topic.TopicId);
        }

        [Test]
        public void Test_007_GetSearchResult()
        {
            IList list = topicDao.GetSearchResult(topic.Title, 0, 10);
            Topic temp = list[0] as Topic;
            Assert.AreEqual(topic.Owner.UserId, temp.Owner.UserId);
            Assert.AreEqual(topic.RecommendedOrder, temp.RecommendedOrder);
        }
    }
}
