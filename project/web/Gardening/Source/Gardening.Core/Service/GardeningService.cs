using System.Collections;
using Gardening.Core.Domain;
using Gardening.Core.Persistence;
using System;

namespace Gardening.Core.Service
{
    public class GardeningService : IGardeningService
    {
        private ITopicDao topicDao = null;
        private IEntryDao entryDao = null;
        private IImgFileDao imgFileDao = null;

        public ITopicDao OwnerDao
        {
            set
            {
                topicDao = value;
            }
        }

        public IEntryDao EntryDao
        {
            set
            {
                entryDao = value;
            }
        }

        public IImgFileDao ImgFileDao
        {
            set
            {
                imgFileDao = value;
            }
        }

        #region IGardeningService Members

        public Topic GetTopic(string topicId)
        {
            Topic result = topicDao.Get(topicId);
            if (result != null)
            {
                result.Avatar = GetTopicAvatar(result.TopicId);
            }

            return result;
        }

        public Topic CreateTopic(Topic topic)
        {
            topic.CreateDateTime = DateTime.Now;
            topic.ModifyDateTime = DateTime.Now;
            Topic result = topicDao.Create(topic);

            if (topic.Avatar != null)
            {
                topic.Avatar.ParentId = result.TopicId;
                result.Avatar = imgFileDao.Create(topic.Avatar);
            }
            return result;
        }

        public Topic UpdateTopic(Topic topic)
        {
            topic.ModifyDateTime = DateTime.Now;
            Topic result = topicDao.Update(topic);

            if (topic.Avatar != null)
            {
                if (topic.Avatar.FileId == null)
                {
                    result.Avatar = imgFileDao.Create(topic.Avatar);
                }
                else
                {
                    result.Avatar = imgFileDao.Update(topic.Avatar);
                }
            }

            return result;
        }

        public IList GetAllTopics()
        {
            return HandleEachTopicAvatar(topicDao.GetAll());
        }

        public IList GetTopicsByOwner(string ownerId)
        {
            return HandleEachTopicAvatar(topicDao.GetByOwner(ownerId));
        }

        public void DeleteTopic(string topicId)
        {
            IList entryList = entryDao.GetByTopic(topicId);
            foreach (Entry e in entryList)
            {
                DeleteEntry(e.EntryId);
            }

            topicDao.Delete(topicId);
        }

        public Entry GetEntry(string entryId)
        {
            if (entryId.Trim() == "")
            {
                return null;
            }

            Entry e = entryDao.Get(entryId);
            e.Files = imgFileDao.GetByParent(entryId);

            return e;
        }

        public IList GetEntriesByTopic(string topicId)
        {
            if (topicId.Trim() == "")
            {
                return null;
            }

            IList list = entryDao.GetByTopic(topicId);

            IList result = new ArrayList();

            foreach (Entry e in list)
            {
                e.Files = imgFileDao.GetByParent(e.EntryId);
                result.Add(e);
            }

            return result;
        }

        public Entry CreateEntry(Entry entry)
        {
            if (string.IsNullOrEmpty(entry.TopicId))
            {
                throw new Exception("invalid topicId");
            }

            entry.CreateDateTime = DateTime.Now;
            entry.ModifyDateTime = DateTime.Now;

            foreach (ImgFile sf in entry.Files)
            {
                sf.ParentId = entry.EntryId;
                sf.FileId = Utility.GetGuid();

                imgFileDao.Create(sf);
            }

            entryDao.Create(entry);

            return GetEntry(entry.EntryId);
        }

        public Entry UpdateEntry(Entry entry)
        {
            entry.ModifyDateTime = DateTime.Now;

            foreach (ImgFile sf in entry.Files)
            {
                imgFileDao.Update(sf);
            }

            entryDao.Update(entry);

            return GetEntry(entry.EntryId);
        }

        public void DeleteEntry(string entryId)
        {
            Entry entry = GetEntry(entryId);

            if (entry.Files != null)
            {
                foreach (ImgFile sf in entry.Files)
                {
                    imgFileDao.Delete(sf.FileId);
                }
            }

            entryDao.Delete(entryId);
        }

        public void DeleteEntryByTopic(string topicId)
        {
            IList list = entryDao.GetByTopic(topicId);

            foreach (Entry e in list)
            {
                DeleteEntry(e.EntryId);
            }
        }

        public void DeleteImgFile(string fileId)
        {
            imgFileDao.Delete(fileId);
        }

        public IList GetAllEntries()
        {
            IList list = entryDao.GetAll();
            IList result = new ArrayList();

            foreach (Entry e in list)
            {
                e.Files = imgFileDao.GetByParent(e.EntryId);
                result.Add(e);
            }

            return result;
        }

        public IList GetTopicsByRecommendedOrder(int topN)
        {
            return HandleEachTopicAvatar(topicDao.GetByRecommendedOrder(topN));
        }

        #endregion

        private IList HandleEachTopicAvatar(IList topicList)
        {
            ArrayList result = new ArrayList();

            foreach (Topic topic in topicList)
            {
                topic.Avatar = GetTopicAvatar(topic.TopicId);
                result.Add(topic);
            }

            return result;
        }

        private ImgFile GetTopicAvatar(string topicId)
        {
            IList iFile = imgFileDao.GetByParent(topicId);
            if (iFile.Count > 0)
            {
                return iFile[0] as ImgFile;
            }
            else
            {
                return null;
            }
        }

        public IList GetTopicsSearchResult(string keyword, int pageIndex, int pageSize)
        {
            return HandleEachTopicAvatar(topicDao.GetSearchResult(keyword, pageIndex, pageSize));
        }
    }
}
