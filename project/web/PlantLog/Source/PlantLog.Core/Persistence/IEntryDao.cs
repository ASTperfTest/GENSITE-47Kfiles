using System.Collections;
using PlantLog.Core.Domain;

namespace PlantLog.Core.Persistence
{
    public interface IEntryDao
    {
        Entry Get(string entryId);
        IList GetByOwner(string ownerId);
        IList GetAll();
        Entry Create(Entry entry);
        Entry Update(Entry entry);
        void Delete(string entryId);
        void DeleteByOwner(string ownerId);
    }
}
