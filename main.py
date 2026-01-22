import sys
from db.models import Base
from db.config import engine
from updater import run_sap_update, run_sap_update_with_job

Base.metadata.create_all(engine)

if __name__ == "__main__":
    # job_id = int(sys.argv[1])
    # run_sap_update_with_job(job_id)
    run_sap_update()