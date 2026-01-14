from datetime import datetime
from typing import Optional
from sqlalchemy.orm import Session
from db.config import get_session
from db.models import UpdateJob


class UpdateJobRepository:

    def __init__(self, session: Optional[Session] = None):
        self.session = session or get_session()
        self._close = session is None

    def exists_running_job(self) -> bool:
        return (
            self.session.query(UpdateJob)
            .filter(UpdateJob.status == "RUNNING")
            .count()
            > 0
        )
    def create_running_job(self) -> UpdateJob:
        job = UpdateJob(
            status="RUNNING",
            started_at=datetime.now()
        )

        self.session.add(job)
        self.session.commit()
        self.session.refresh(job)

        return job
    
    def mark_success(self, job_id: int):
        job = self.session.get(UpdateJob, job_id)

        job.status = "SUCCESS"
        job.finished_at = datetime.now()

        self.session.commit()
   
    def mark_error(self, job_id: int, error_message: str):
        job = self.session.get(UpdateJob, job_id)

        job.status = "ERROR"
        job.error_message = error_message
        job.finished_at = datetime.now()

        self.session.commit()

    def get_last_job(self) -> UpdateJob | None:
        return (
            self.session.query(UpdateJob)
            .order_by(UpdateJob.id.desc())
            .first()
        )
    
    def __del__(self):
        if self._close:
            self.session.close()

