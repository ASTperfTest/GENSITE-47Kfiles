using PlantLog.Core.Domain;
using System.Collections;

namespace PlantLog.Core.Persistence
{
    public interface IVoteRecordDao
    {
        IList GetByUser(string userId);
        IList GetVoteRecordByVoter(string voterId);
        IList GetAll();
        void Create(VoteRecord voteRecord);
        void Update(VoteRecord voteRecord);
    }
}
