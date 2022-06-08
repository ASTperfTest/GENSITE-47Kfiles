using System.Collections;
using Gardening.Core.Domain;

namespace Gardening.Core.Persistence
{
    public interface IEntryDao
    {
        Entry Get(string entryId);
        IList GetByTopic(string topicId);
        IList GetAll();
        Entry Create(Entry entry);
        Entry Update(Entry entry);
        void Delete(string entryId);
        void DeleteByTopic(string topicId);
    }
}
