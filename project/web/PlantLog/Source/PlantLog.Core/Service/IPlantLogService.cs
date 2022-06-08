using System.Collections;
using PlantLog.Core.Domain;

namespace PlantLog.Core.Service
{
    public interface IPlantLogService
    {
        Owner GetOwner(string ownerId);
        Owner CreateOwner(Owner owner);
        Owner UpdateOwner(Owner owner);
        IList GetAllOwner();
        IList GetAllEntry();
        void DeleteOwner(string ownerId);
        Entry GetEntry(string entryId);
        IList GetEntryByOwner(string ownerId);
        Entry CreateEntry(Entry entry);
        Entry UpdateEntry(Entry entry);
        void DeleteEntry(string entryId);
        void DeleteEntryByOwner(string ownerId);
        void DeleteImgFile(string fileId);
        IList GetAllVoteRecord();
        IList GetVoteRecordByUser(string userId);
        IList GetVoteRecordByVoter(string voterId);
        void CreateVoteRecord(VoteRecord voteRecord);
        void UpdateVoteRecord(VoteRecord voteRecord);
    }
}
