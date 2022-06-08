using System.Collections;
using Gardening.Core.Domain;

namespace Gardening.Core.Persistence
{
    public interface ITopicDao
    {
        IList GetByOwner(string ownerId);
        Topic Get(string topicId);
        IList GetAll();
        Topic Create(Topic topic);
        Topic Update(Topic topic);
        void Delete(string topicId);
        IList GetByRecommendedOrder(int topN);
		IList GetSearchResult(string keyword,int pageIndex,int pageSize);
    }
}
