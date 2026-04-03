from pydantic import BaseModel, EmailStr, HttpUrl
from typing import Optional


class Header(BaseModel):
    full_name: str
    email: str
    number: str
    location: str
    linkdin: str
    github: str


class Education(BaseModel):
    school: str
    study: str
    expected: str
    GPA: bool
    award: str
    relevant_courses: list[str]


class Experience(BaseModel):
    job_title: str
    timeline: str
    organization: str
    location: str
    responsiblities: list[str]


class Project(BaseModel):
    name: str
    description: str
    link: str
    bullet_points: list[str]


class TechnicalSkills(BaseModel):
    languages: list[str]
    frameworks_and_libraries: list[str]
    ai_ml_capabilities: list[str]
    soft_skills: list[str]


class Resume(BaseModel):
    header: Header
    summary: str
    education: Education
    experience: list[Experience]
    projects: list[Project]
    technical_skills: TechnicalSkills