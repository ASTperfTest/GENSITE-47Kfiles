using System.Collections;
using Gardening.Core.Domain;

namespace Gardening.Core.Service
{
    public interface IGardeningService
    {
        Topic GetTopic(string topicId);
        IList GetTopicsByOwner(string ownerId);
        Topic CreateTopic(Topic topic);
        Topic UpdateTopic(Topic topic);
        IList GetAllTopics();
        IList GetTopicsByRecommendedOrder(int topN);
        IList GetAllEntries();
        void DeleteTopic(string topicId);
        Entry GetEntry(string entryId);
        IList GetEntriesByTopic(string topicId);
        Entry CreateEntry(Entry entry);
        Entry UpdateEntry(Entry entry);
        void DeleteEntry(string entryId);
        void DeleteEntryByTopic(string topicId);
        void DeleteImgFile(string fileId);
		IList GetTopicsSearchResult(string keyword,int pageIndex,int pageSize);
    }
}
