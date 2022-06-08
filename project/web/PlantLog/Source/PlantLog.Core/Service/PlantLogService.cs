using System.Collections;
using PlantLog.Core.Domain;
using PlantLog.Core.Persistence;
using System;

namespace PlantLog.Core.Service
{
    public class PlantLogService : IPlantLogService
    {
        private IOwnerDao ownerDao = null;
        private IEntryDao entryDao = null;
        private IImgFileDao imgFileDao = null;
        private IVoteRecordDao voteRecordDao = null;

        public IOwnerDao OwnerDao
        {
            set
            {
                ownerDao = value;
            }
        }

        public IEntryDao EntryDao
        {
            set
            {
                entryDao = value;
            }
        }

        public IImgFileDao SourceFileDao
        {
            set
            {
                imgFileDao = value;
            }
        }

        public IVoteRecordDao VoteRecordDao
        {
            set
            {
                voteRecordDao = value;
            }
        }

        #region IPlantLogService Members

        public Owner GetOwner(string ownerId)
        {
            Owner result = ownerDao.Get(ownerId);
            if (result != null)
            {
                IList iFile = imgFileDao.GetByEntry(ownerId);
                if (iFile.Count > 0)
                {
                    result.Avatar = iFile[0] as ImgFile;
                }
                else
                {
                    result.Avatar = null;
                }
                result.VoteRecords = voteRecordDao.GetByUser(ownerId);
            }

            return result;
        }

        public Owner CreateOwner(Owner owner)
        {
            Owner user = ownerDao.Create(owner);

            return user;
        }

        public Owner UpdateOwner(Owner owner)
        {
            Owner user = ownerDao.Update(owner);

            if (owner.Avatar != null)
            {
                if (owner.Avatar.FileId == null)
                {
                    owner.Avatar.FileId = Utility.GetGuid();
                    user.Avatar = imgFileDao.Create(owner.Avatar);
                }
                else
                {
                    user.Avatar = imgFileDao.Update(owner.Avatar);
                }
            }

            return user;
        }

        public IList GetAllOwner()
        {
            IList temp = ownerDao.GetAll();
            ArrayList result = new ArrayList();

            foreach (Owner o in temp)
            {
                IList iFile = imgFileDao.GetByEntry(o.OwnerId);
                if (iFile.Count > 0)
                {
                    o.Avatar = iFile[0] as ImgFile;
                }
                else
                {
                    o.Avatar = null;
                }
                o.VoteRecords = voteRecordDao.GetByUser(o.OwnerId);
                result.Add(o);
            }

            return result;
        }

        public void DeleteOwner(string ownerId)
        {
            entryDao.DeleteByOwner(ownerId);

            ownerDao.Delete(ownerId);
        }

        public Entry GetEntry(string entryId)
        {
            if (entryId.Trim() == "")
            {
                return null;
            }

            Entry e = entryDao.Get(entryId);
            e.Files = imgFileDao.GetByEntry(entryId);

            return e;
        }

        public IList GetEntryByOwner(string ownerId)
        {
            if (ownerId.Trim() == "")
            {
                return null;
            }

            IList list = entryDao.GetByOwner(ownerId);

            IList result = new ArrayList();

            foreach (Entry e in list)
            {
                e.Files = imgFileDao.GetByEntry(e.EntryId);
                result.Add(e);
            }

            return result;
        }

        public Entry CreateEntry(Entry entry)
        {
            foreach (ImgFile sf in entry.Files)
            {
                sf.EntryId = entry.EntryId;
                sf.FileId = Utility.GetGuid();

                imgFileDao.Create(sf);
            }

            entryDao.Create(entry);

            return GetEntry(entry.EntryId);
        }

        public Entry UpdateEntry(Entry entry)
        {
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

        public void DeleteEntryByOwner(string ownerId)
        {
            IList list = entryDao.GetByOwner(ownerId);

            foreach (Entry e in list)
            {
                DeleteEntry(e.EntryId);
            }
        }

        public void DeleteImgFile(string fileId)
        {
            imgFileDao.Delete(fileId);
        }

        public IList GetAllVoteRecord()
        {
            return voteRecordDao.GetAll();
        }

        public IList GetVoteRecordByUser(string userId)
        {
            return voteRecordDao.GetByUser(userId);
        }

        public IList GetVoteRecordByVoter(string voterId)
        {
            return voteRecordDao.GetVoteRecordByVoter(voterId);
        }

        public void CreateVoteRecord(VoteRecord voteRecord)
        {
            voteRecordDao.Create(voteRecord);
        }

        public void UpdateVoteRecord(VoteRecord voteRecord)
        {
            voteRecordDao.Update(voteRecord);
        }

        public IList GetAllEntry()
        {
            IList list = entryDao.GetAll();
            IList result = new ArrayList();

            foreach (Entry e in list)
            {
                e.Files = imgFileDao.GetByEntry(e.EntryId);
                result.Add(e);
            }

            return result;
        }

        #endregion
    }
}
