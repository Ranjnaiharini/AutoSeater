from fastapi import FastAPI, APIRouter, HTTPException, Depends, status
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from fastapi.responses import StreamingResponse
from dotenv import load_dotenv
from starlette.middleware.cors import CORSMiddleware
from motor.motor_asyncio import AsyncIOMotorClient
import os
import logging
from pathlib import Path
from pydantic import BaseModel, Field, ConfigDict, EmailStr
from typing import List, Optional, Dict, Any
import uuid
from datetime import datetime, timezone, timedelta
from passlib.context import CryptContext
import jwt
import io
import pandas as pd
from enum import Enum

ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

# MongoDB connection
mongo_url = os.environ['MONGO_URL']
client = AsyncIOMotorClient(mongo_url)
db = client[os.environ['DB_NAME']]

# Security
pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
security = HTTPBearer()
SECRET_KEY = os.environ.get('JWT_SECRET_KEY', 'autoseater-secret-key-change-in-production')
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 1440  # 24 hours

# Create the main app
app = FastAPI()
api_router = APIRouter(prefix="/api")


# Simple API root/health endpoint to avoid 404s from frontend callers to /api
@app.get("/api")
async def api_root():
    return {"ok": True, "message": "AutoSeater API running"}

# Enums
class UserRole(str, Enum):
    ADMIN = "admin"
    INVIGILATOR = "invigilator"

class ExamType(str, Enum):
    CAT = "CAT"
    SEMESTER = "Semester"

class SeatingMode(str, Enum):
    TWO_PER_DESK = "two_per_desk"  # CAT mode
    ONE_PER_DESK = "one_per_desk"  # Semester mode

# Models
class User(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    username: str
    email: EmailStr
    role: UserRole
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class UserCreate(BaseModel):
    username: str
    email: EmailStr
    password: str
    role: UserRole = UserRole.ADMIN

class UserLogin(BaseModel):
    username: str
    password: str

class Token(BaseModel):
    access_token: str
    token_type: str
    user: User

class Student(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    roll_number: str
    name: str
    department: str
    subjects: List[str]
    email: Optional[str] = None
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class StudentCreate(BaseModel):
    roll_number: str
    name: str
    department: str
    subjects: List[str]
    email: Optional[str] = None

class StudentUpdate(BaseModel):
    name: Optional[str] = None
    department: Optional[str] = None
    subjects: Optional[List[str]] = None
    email: Optional[str] = None

class Department(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    name: str
    code: str
    subjects: List[str]
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class DepartmentCreate(BaseModel):
    name: str
    code: str
    subjects: List[str]

class Room(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    name: str
    capacity: int
    desk_count: int
    rows: int
    columns: int
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class RoomCreate(BaseModel):
    name: str
    capacity: int
    desk_count: int
    rows: int
    columns: int

class Exam(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    exam_name: str
    exam_type: ExamType
    date: str
    time: str
    departments: List[str]
    subjects: List[str]
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))
    created_by: str

class ExamCreate(BaseModel):
    exam_name: str
    exam_type: ExamType
    date: str
    time: str
    departments: List[str]
    subjects: List[str]

class DeskAssignment(BaseModel):
    desk_number: int
    left_student: Optional[str] = None  # roll_number
    right_student: Optional[str] = None  # roll_number
    row: int
    col: int

class SeatingPlan(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    exam_id: str
    room_id: str
    seating_mode: SeatingMode
    desk_assignments: List[DeskAssignment]
    total_students: int
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class SeatingGenerateRequest(BaseModel):
    exam_id: str
    room_ids: List[str]
    seating_mode: SeatingMode

# Helper functions
def hash_password(password: str) -> str:
    return pwd_context.hash(password)

def verify_password(plain_password: str, hashed_password: str) -> bool:
    return pwd_context.verify(plain_password, hashed_password)

def create_access_token(data: dict) -> str:
    to_encode = data.copy()
    expire = datetime.now(timezone.utc) + timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    to_encode.update({"exp": expire})
    encoded_jwt = jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)
    return encoded_jwt

async def get_current_user(credentials: HTTPAuthorizationCredentials = Depends(security)) -> User:
    try:
        token = credentials.credentials
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        user_id: str = payload.get("sub")
        if user_id is None:
            raise HTTPException(status_code=401, detail="Invalid authentication credentials")
    except jwt.PyJWTError:
        raise HTTPException(status_code=401, detail="Invalid authentication credentials")
    
    user_doc = await db.users.find_one({"id": user_id}, {"_id": 0})
    if user_doc is None:
        raise HTTPException(status_code=401, detail="User not found")
    
    # Convert ISO string to datetime if needed
    if isinstance(user_doc.get('created_at'), str):
        user_doc['created_at'] = datetime.fromisoformat(user_doc['created_at'])
    
    return User(**user_doc)

async def get_admin_user(current_user: User = Depends(get_current_user)) -> User:
    if current_user.role != UserRole.ADMIN:
        raise HTTPException(status_code=403, detail="Admin access required")
    return current_user

# Authentication Routes
@api_router.post("/auth/register", response_model=User)
async def register(user_data: UserCreate):
    # Check if username exists
    existing_user = await db.users.find_one({"username": user_data.username})
    if existing_user:
        raise HTTPException(status_code=400, detail="Username already exists")
    
    # Check if email exists
    existing_email = await db.users.find_one({"email": user_data.email})
    if existing_email:
        raise HTTPException(status_code=400, detail="Email already exists")
    
    # Hash password
    hashed_password = hash_password(user_data.password)
    
    # Create user
    user = User(
        username=user_data.username,
        email=user_data.email,
        role=user_data.role
    )
    
    doc = user.model_dump()
    doc['password'] = hashed_password
    doc['created_at'] = doc['created_at'].isoformat()
    
    await db.users.insert_one(doc)
    return user

@api_router.post("/auth/login", response_model=Token)
async def login(credentials: UserLogin):
    user_doc = await db.users.find_one({"username": credentials.username}, {"_id": 0})
    if not user_doc:
        raise HTTPException(status_code=401, detail="Invalid username or password")
    
    if not verify_password(credentials.password, user_doc['password']):
        raise HTTPException(status_code=401, detail="Invalid username or password")
    
    # Convert ISO string to datetime
    if isinstance(user_doc.get('created_at'), str):
        user_doc['created_at'] = datetime.fromisoformat(user_doc['created_at'])
    
    user = User(**{k: v for k, v in user_doc.items() if k != 'password'})
    
    access_token = create_access_token(data={"sub": user.id})
    
    return Token(
        access_token=access_token,
        token_type="bearer",
        user=user
    )

@api_router.get("/auth/me", response_model=User)
async def get_me(current_user: User = Depends(get_current_user)):
    return current_user

# Student Routes
@api_router.post("/students", response_model=Student)
async def create_student(student_data: StudentCreate, current_user: User = Depends(get_admin_user)):
    # Check if roll number exists
    existing = await db.students.find_one({"roll_number": student_data.roll_number})
    if existing:
        raise HTTPException(status_code=400, detail="Roll number already exists")
    
    student = Student(**student_data.model_dump())
    doc = student.model_dump()
    doc['created_at'] = doc['created_at'].isoformat()
    
    await db.students.insert_one(doc)
    return student

@api_router.post("/students/bulk", response_model=Dict[str, Any])
async def create_students_bulk(students_data: List[StudentCreate], current_user: User = Depends(get_admin_user)):
    created = 0
    skipped = 0
    errors = []
    
    for student_data in students_data:
        try:
            existing = await db.students.find_one({"roll_number": student_data.roll_number})
            if existing:
                skipped += 1
                errors.append(f"Roll number {student_data.roll_number} already exists")
                continue
            
            student = Student(**student_data.model_dump())
            doc = student.model_dump()
            doc['created_at'] = doc['created_at'].isoformat()
            
            await db.students.insert_one(doc)
            created += 1
        except Exception as e:
            errors.append(f"Error creating student {student_data.roll_number}: {str(e)}")
            skipped += 1
    
    return {
        "created": created,
        "skipped": skipped,
        "errors": errors
    }

@api_router.get("/students", response_model=List[Student])
async def get_students(current_user: User = Depends(get_current_user)):
    students = await db.students.find({}, {"_id": 0}).to_list(10000)
    for student in students:
        if isinstance(student.get('created_at'), str):
            student['created_at'] = datetime.fromisoformat(student['created_at'])
    return students

@api_router.get("/students/{student_id}", response_model=Student)
async def get_student(student_id: str, current_user: User = Depends(get_current_user)):
    student = await db.students.find_one({"id": student_id}, {"_id": 0})
    if not student:
        raise HTTPException(status_code=404, detail="Student not found")
    if isinstance(student.get('created_at'), str):
        student['created_at'] = datetime.fromisoformat(student['created_at'])
    return Student(**student)

@api_router.put("/students/{student_id}", response_model=Student)
async def update_student(student_id: str, student_data: StudentUpdate, current_user: User = Depends(get_admin_user)):
    student = await db.students.find_one({"id": student_id}, {"_id": 0})
    if not student:
        raise HTTPException(status_code=404, detail="Student not found")
    
    update_data = {k: v for k, v in student_data.model_dump().items() if v is not None}
    if update_data:
        await db.students.update_one({"id": student_id}, {"$set": update_data})
        student.update(update_data)
    
    if isinstance(student.get('created_at'), str):
        student['created_at'] = datetime.fromisoformat(student['created_at'])
    return Student(**student)

@api_router.delete("/students/{student_id}")
async def delete_student(student_id: str, current_user: User = Depends(get_admin_user)):
    result = await db.students.delete_one({"id": student_id})
    if result.deleted_count == 0:
        raise HTTPException(status_code=404, detail="Student not found")
    return {"message": "Student deleted successfully"}

# Department Routes
@api_router.post("/departments", response_model=Department)
async def create_department(dept_data: DepartmentCreate, current_user: User = Depends(get_admin_user)):
    existing = await db.departments.find_one({"code": dept_data.code})
    if existing:
        raise HTTPException(status_code=400, detail="Department code already exists")
    
    dept = Department(**dept_data.model_dump())
    doc = dept.model_dump()
    doc['created_at'] = doc['created_at'].isoformat()
    
    await db.departments.insert_one(doc)
    return dept

@api_router.get("/departments", response_model=List[Department])
async def get_departments(current_user: User = Depends(get_current_user)):
    depts = await db.departments.find({}, {"_id": 0}).to_list(1000)
    for dept in depts:
        if isinstance(dept.get('created_at'), str):
            dept['created_at'] = datetime.fromisoformat(dept['created_at'])
    return depts

@api_router.delete("/departments/{dept_id}")
async def delete_department(dept_id: str, current_user: User = Depends(get_admin_user)):
    result = await db.departments.delete_one({"id": dept_id})
    if result.deleted_count == 0:
        raise HTTPException(status_code=404, detail="Department not found")
    return {"message": "Department deleted successfully"}

# Room Routes
@api_router.post("/rooms", response_model=Room)
async def create_room(room_data: RoomCreate, current_user: User = Depends(get_admin_user)):
    room = Room(**room_data.model_dump())
    doc = room.model_dump()
    doc['created_at'] = doc['created_at'].isoformat()
    
    await db.rooms.insert_one(doc)
    return room

@api_router.get("/rooms", response_model=List[Room])
async def get_rooms(current_user: User = Depends(get_current_user)):
    rooms = await db.rooms.find({}, {"_id": 0}).to_list(1000)
    for room in rooms:
        if isinstance(room.get('created_at'), str):
            room['created_at'] = datetime.fromisoformat(room['created_at'])
    return rooms

@api_router.delete("/rooms/{room_id}")
async def delete_room(room_id: str, current_user: User = Depends(get_admin_user)):
    result = await db.rooms.delete_one({"id": room_id})
    if result.deleted_count == 0:
        raise HTTPException(status_code=404, detail="Room not found")
    return {"message": "Room deleted successfully"}

# Exam Routes
@api_router.post("/exams", response_model=Exam)
async def create_exam(exam_data: ExamCreate, current_user: User = Depends(get_admin_user)):
    exam = Exam(**exam_data.model_dump(), created_by=current_user.id)
    doc = exam.model_dump()
    doc['created_at'] = doc['created_at'].isoformat()
    
    await db.exams.insert_one(doc)
    return exam

@api_router.get("/exams", response_model=List[Exam])
async def get_exams(current_user: User = Depends(get_current_user)):
    exams = await db.exams.find({}, {"_id": 0}).to_list(1000)
    for exam in exams:
        if isinstance(exam.get('created_at'), str):
            exam['created_at'] = datetime.fromisoformat(exam['created_at'])
    return exams

@api_router.get("/exams/{exam_id}", response_model=Exam)
async def get_exam(exam_id: str, current_user: User = Depends(get_current_user)):
    exam = await db.exams.find_one({"id": exam_id}, {"_id": 0})
    if not exam:
        raise HTTPException(status_code=404, detail="Exam not found")
    if isinstance(exam.get('created_at'), str):
        exam['created_at'] = datetime.fromisoformat(exam['created_at'])
    return Exam(**exam)

@api_router.delete("/exams/{exam_id}")
async def delete_exam(exam_id: str, current_user: User = Depends(get_admin_user)):
    # Delete associated seating plans
    await db.seating_plans.delete_many({"exam_id": exam_id})
    
    result = await db.exams.delete_one({"id": exam_id})
    if result.deleted_count == 0:
        raise HTTPException(status_code=404, detail="Exam not found")
    return {"message": "Exam and associated seating plans deleted successfully"}

# Seating Generation
@api_router.post("/seating/generate")
async def generate_seating(request: SeatingGenerateRequest, current_user: User = Depends(get_admin_user)):
    # Get exam
    exam = await db.exams.find_one({"id": request.exam_id}, {"_id": 0})
    if not exam:
        raise HTTPException(status_code=404, detail="Exam not found")
    
    # Log exam filter info for debugging
    logger.info(f"Seating generation requested for exam_id={request.exam_id} departments={exam.get('departments')} subjects={exam.get('subjects')}")

    # Get eligible students
    eligible_students = await db.students.find({
        "department": {"$in": exam['departments']},
        "subjects": {"$in": exam['subjects']}
    }, {"_id": 0}).to_list(10000)
    
    logger.info(f"Found {len(eligible_students)} eligible students for exam_id={request.exam_id}")

    if not eligible_students:
        # Provide a helpful log entry to assist debugging in dev environments
        logger.warning(f"No eligible students found for exam_id={request.exam_id} using filter departments={exam.get('departments')} subjects={exam.get('subjects')}")
        raise HTTPException(status_code=400, detail="No eligible students found")
    
    # Get rooms
    rooms = await db.rooms.find({"id": {"$in": request.room_ids}}, {"_id": 0}).to_list(100)
    if not rooms:
        raise HTTPException(status_code=404, detail="No rooms found")
    
    # Sort students by roll number
    eligible_students.sort(key=lambda s: s['roll_number'])
    
    # Delete existing seating plans for this exam
    await db.seating_plans.delete_many({"exam_id": request.exam_id})
    
    seating_plans = []
    student_index = 0
    
    for room in rooms:
        if student_index >= len(eligible_students):
            break
        
        desk_assignments = []
        desk_number = 1
        
        if request.seating_mode == SeatingMode.TWO_PER_DESK:
            # CAT mode: 2 students per desk from different subjects
            # Group students by subject
            subject_groups = {}
            for student in eligible_students[student_index:]:
                for subject in student['subjects']:
                    if subject in exam['subjects']:
                        if subject not in subject_groups:
                            subject_groups[subject] = []
                        subject_groups[subject].append(student)
                        break
            
            # Create pairs from different subjects
            subject_lists = list(subject_groups.values())
            max_desks = room['desk_count']
            
            for row in range(room['rows']):
                for col in range(room['columns']):
                    if desk_number > max_desks:
                        break
                    
                    left_student = None
                    right_student = None
                    
                    # Try to assign students from different subjects
                    if len(subject_lists) >= 2:
                        if subject_lists[0]:
                            left_student = subject_lists[0].pop(0)['roll_number']
                        if subject_lists[1]:
                            right_student = subject_lists[1].pop(0)['roll_number']
                    elif len(subject_lists) == 1 and subject_lists[0]:
                        if len(subject_lists[0]) >= 2:
                            left_student = subject_lists[0].pop(0)['roll_number']
                            right_student = subject_lists[0].pop(0)['roll_number']
                        elif len(subject_lists[0]) == 1:
                            left_student = subject_lists[0].pop(0)['roll_number']
                    
                    if left_student or right_student:
                        desk_assignments.append(DeskAssignment(
                            desk_number=desk_number,
                            left_student=left_student,
                            right_student=right_student,
                            row=row,
                            col=col
                        ))
                    
                    desk_number += 1
                    
                    if not any(subject_lists):
                        break
                if not any(subject_lists):
                    break
            
            # Count how many students were actually assigned to desks in this room
            assigned_in_room = 0
            for d in desk_assignments:
                if d.left_student:
                    assigned_in_room += 1
                if d.right_student:
                    assigned_in_room += 1

            student_index += assigned_in_room
        
        else:  # ONE_PER_DESK (Semester mode)
            # 1 student per desk, sequential roll order
            # Vertical pairing: desk 1 with 25, 2 with 26, etc.
            students_for_room = eligible_students[student_index:student_index + room['desk_count']]
            
            # Calculate vertical pairing
            half_desks = room['desk_count'] // 2
            
            for row in range(room['rows']):
                for col in range(room['columns']):
                    if desk_number > room['desk_count'] or student_index >= len(eligible_students):
                        break
                    
                    student = eligible_students[student_index] if student_index < len(eligible_students) else None
                    
                    if student:
                        desk_assignments.append(DeskAssignment(
                            desk_number=desk_number,
                            left_student=student['roll_number'],
                            right_student=None,
                            row=row,
                            col=col
                        ))
                        student_index += 1
                    
                    desk_number += 1
        
        # Create seating plan
        total_students = sum(1 for d in desk_assignments if d.left_student) + sum(1 for d in desk_assignments if d.right_student)
        
        seating_plan = SeatingPlan(
            exam_id=request.exam_id,
            room_id=room['id'],
            seating_mode=request.seating_mode,
            desk_assignments=[d.model_dump() for d in desk_assignments],
            total_students=total_students
        )
        
        doc = seating_plan.model_dump()
        doc['created_at'] = doc['created_at'].isoformat()
        
        await db.seating_plans.insert_one(doc)
        seating_plans.append(seating_plan)
    
    return {
        "message": "Seating plans generated successfully",
        "plans_created": len(seating_plans),
        "total_students_assigned": student_index,
        "total_eligible_students": len(eligible_students)
    }

@api_router.get("/seating/exam/{exam_id}")
async def get_seating_plans(exam_id: str, current_user: User = Depends(get_current_user)):
    plans = await db.seating_plans.find({"exam_id": exam_id}, {"_id": 0}).to_list(100)
    for plan in plans:
        if isinstance(plan.get('created_at'), str):
            plan['created_at'] = datetime.fromisoformat(plan['created_at'])
    
    # Get room details
    for plan in plans:
        room = await db.rooms.find_one({"id": plan['room_id']}, {"_id": 0})
        plan['room_details'] = room
    
    return plans

@api_router.get("/seating/export/{exam_id}")
async def export_seating_excel(exam_id: str, current_user: User = Depends(get_current_user)):
    # Get exam
    exam = await db.exams.find_one({"id": exam_id}, {"_id": 0})
    if not exam:
        raise HTTPException(status_code=404, detail="Exam not found")
    
    # Get seating plans
    plans = await db.seating_plans.find({"exam_id": exam_id}, {"_id": 0}).to_list(100)
    if not plans:
        raise HTTPException(status_code=404, detail="No seating plans found")
    
    # Create Excel file
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    
    for plan in plans:
        # Get room details
        room = await db.rooms.find_one({"id": plan['room_id']}, {"_id": 0})
        
        # Prepare data
        data = []
        for desk in plan['desk_assignments']:
            data.append({
                'Desk Number': desk['desk_number'],
                'Row': desk['row'] + 1,
                'Column': desk['col'] + 1,
                'Left Student': desk.get('left_student', 'Empty'),
                'Right Student': desk.get('right_student', 'Empty')
            })
        
        df = pd.DataFrame(data)
        sheet_name = room['name'][:31]  # Excel sheet name limit
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    writer.close()
    output.seek(0)
    
    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={"Content-Disposition": f"attachment; filename=seating_plan_{exam['exam_name']}.xlsx"}
    )

# Dashboard Stats
@api_router.get("/dashboard/stats")
async def get_dashboard_stats(current_user: User = Depends(get_current_user)):
    total_students = await db.students.count_documents({})
    total_exams = await db.exams.count_documents({})
    total_rooms = await db.rooms.count_documents({})
    total_departments = await db.departments.count_documents({})
    
    # Get recent exams
    recent_exams = await db.exams.find({}, {"_id": 0}).sort("created_at", -1).limit(5).to_list(5)
    
    return {
        "total_students": total_students,
        "total_exams": total_exams,
        "total_rooms": total_rooms,
        "total_departments": total_departments,
        "recent_exams": recent_exams
    }

# Include router
app.include_router(api_router)

app.add_middleware(
    CORSMiddleware,
    allow_credentials=True,
    allow_origins=os.environ.get('CORS_ORIGINS', '*').split(','),
    allow_methods=["*"],
    allow_headers=["*"],
)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@app.on_event("shutdown")
async def shutdown_db_client():
    client.close()