import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import sys
import json
import base64
import re
import pandas as pd
from datetime import datetime
import tempfile
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import requests
import glob
import csv
from pathlib import Path

# ========================================
# CONFIGURATION
# ========================================
OCR_API_KEY = "K86485627588957"
SENDER_EMAIL = "nitesh.t.mulam2004@gmail.com"
APP_PASSWORD = "gxdd zdyh gfym mlcq"
OUTPUT_DIR = "extracted_pdfs"

class SubjectData:
    """Class to hold subject information"""
    
    def __init__(self, name, master_pdf_path, student_pdfs=None):
        self.name = name
        self.master_pdf_path = master_pdf_path
        self.student_pdfs = student_pdfs if student_pdfs else []
        self.results = None
        self.results_file = None
        
    def to_dict(self):
        """Convert to dictionary for saving"""
        return {
            'name': self.name,
            'master_pdf': self.master_pdf_path,
            'student_pdfs': self.student_pdfs
        }
    
    @classmethod
    def from_dict(cls, data):
        """Create from dictionary"""
        return cls(
            name=data['name'],
            master_pdf_path=data['master_pdf'],
            student_pdfs=data.get('student_pdfs', [])
        )

class MultiSubjectPDFProcessor:
    """Handles PDF processing including OCR extraction for multiple subjects"""
    
    def __init__(self, api_key):
        self.api_key = api_key
        self.log_messages = []
    
    def log(self, message, widget=None):
        """Log message to console and optionally to GUI widget"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        print(log_entry)
        self.log_messages.append(log_entry)
        
        if widget:
            widget.insert(tk.END, log_entry + "\n")
            widget.see(tk.END)
    
    def extract_pdf_text(self, pdf_path):
        """Extract text from PDF using PyPDF2"""
        text = ''
        try:
            with open(pdf_path, 'rb') as file:
                reader = PdfReader(file)
                for page in reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + '\n'
            return text
        except Exception as e:
            self.log(f"✗ Error reading PDF {pdf_path}: {e}")
            return ''
    
    def split_pdf_into_chunks(self, pdf_path, pages_per_chunk=3):
        """Split PDF into chunks for OCR processing"""
        try:
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)
            
            if total_pages <= pages_per_chunk:
                return [pdf_path]
            
            chunks = []
            temp_dir = tempfile.mkdtemp()
            
            for start_page in range(0, total_pages, pages_per_chunk):
                end_page = min(start_page + pages_per_chunk, total_pages)
                chunk_pages = list(range(start_page, end_page))
                
                if not chunk_pages:
                    continue
                    
                writer = PdfWriter()
                for page_num in chunk_pages:
                    writer.add_page(reader.pages[page_num])
                
                chunk_filename = os.path.join(temp_dir, f"chunk_{start_page//pages_per_chunk + 1}.pdf")
                with open(chunk_filename, 'wb') as output_pdf:
                    writer.write(output_pdf)
                
                chunks.append(chunk_filename)
            
            return chunks
            
        except Exception as e:
            self.log(f"❌ Error splitting PDF: {str(e)}")
            return [pdf_path]
    
    def extract_text_with_ocr(self, pdf_path, log_widget=None):
        """Extract text using OCR.space API with chunking"""
        try:
            self.log(f"🔄 Starting OCR extraction for {os.path.basename(pdf_path)}", log_widget)
            
            # Split PDF into chunks
            chunks = self.split_pdf_into_chunks(pdf_path, pages_per_chunk=3)
            self.log(f"📦 Split into {len(chunks)} chunks", log_widget)
            
            all_text = ""
            
            for idx, chunk_file in enumerate(chunks, 1):
                self.log(f"🔍 Processing chunk {idx}/{len(chunks)}", log_widget)
                
                # Read chunk as binary
                with open(chunk_file, "rb") as file:
                    pdf_bytes = file.read()
                
                # Convert to base64
                pdf_base64 = base64.b64encode(pdf_bytes).decode('utf-8')
                
                # OCR.space API call
                url = "https://api.ocr.space/parse/image"
                data = {
                    "apikey": self.api_key,
                    "base64Image": f"data:application/pdf;base64,{pdf_base64}",
                    "language": "eng",
                    "isOverlayRequired": False,
                    "OCREngine": 2,
                    "scale": True,
                    "isTable": True,
                    "detectOrientation": True,
                    "filetype": "pdf",
                }
                
                response = requests.post(url, data=data, timeout=60)
                result = json.loads(response.text)
                
                if result.get("IsErroredOnProcessing", False):
                    self.log(f"⚠️ OCR error in chunk {idx}", log_widget)
                    continue
                
                # Extract text from all pages
                parsed_results = result.get("ParsedResults", [])
                for page in parsed_results:
                    page_text = page.get("ParsedText", "")
                    if page_text.strip():
                        all_text += page_text + "\n\n"
                
                # Clean up temporary chunk file
                if os.path.exists(chunk_file):
                    os.remove(chunk_file)
                
                # Small delay to avoid rate limiting
                if idx < len(chunks):
                    import time
                    time.sleep(1)
            
            # Clean extracted text
            all_text = re.sub(r'\s+', ' ', all_text)
            all_text = re.sub(r'(Q\d+)', r'\n\1:', all_text)
            
            self.log(f"✅ OCR extraction complete. Characters: {len(all_text)}", log_widget)
            return all_text.strip()
            
        except Exception as e:
            self.log(f"❌ OCR extraction failed: {str(e)}", log_widget)
            return None
    
    def create_searchable_pdf(self, text_content, output_path):
        """Create a PDF from extracted text"""
        try:
            from reportlab.lib.pagesizes import letter
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.enums import TA_LEFT
            from reportlab.lib.units import inch
            
            doc = SimpleDocTemplate(
                output_path, 
                pagesize=letter,
                leftMargin=0.75*inch,
                rightMargin=0.75*inch,
                topMargin=0.75*inch,
                bottomMargin=0.75*inch
            )
            
            styles = getSampleStyleSheet()
            text_style = ParagraphStyle(
                'ExtractedText',
                parent=styles['Normal'],
                fontSize=11,
                leading=14,
                alignment=TA_LEFT,
                wordWrap='CJK'
            )
            
            story = []
            lines = text_content.split('\n')
            
            for line in lines:
                if line.strip():
                    story.append(Paragraph(line.strip(), text_style))
                    story.append(Spacer(1, 6))
            
            if story:
                doc.build(story)
                return True
            return False
            
        except ImportError:
            self.log("⚠️ ReportLab not installed. Cannot create PDF.")
            return False
        except Exception as e:
            self.log(f"❌ Error creating PDF: {str(e)}")
            return False

class EmailSender:
    """Handles email sending functionality for multiple subjects"""
    
    def __init__(self, sender_email, app_password):
        self.sender_email = sender_email
        self.app_password = app_password
        self.sent_emails_log = []
    
    def test_connection(self):
        """Test email connection"""
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(self.sender_email, self.app_password)
            return True, "Email connection successful"
        except Exception as e:
            return False, f"Email connection failed: {str(e)}"
    
    def send_results_email(self, student_data, subject_results, results_file=None):
        """Send results email to student with multiple subjects"""
        try:
            student_name = student_data['Name']
            student_email = student_data['Email']
            roll_no = student_data['Roll No']
            
            # Validate email
            if not student_email or student_email == f"student_{student_data.get('id', 0)}@example.com":
                return False, f"No valid email found for {student_name}"
            
            # Create email message
            message = MIMEMultipart()
            message["From"] = self.sender_email
            message["To"] = student_email
            message["Subject"] = f"📊 Exam Results - {student_name}"
            
            # Create subject-wise table
            subject_table = ""
            for subject_name, scores in subject_results.items():
                total = scores.get('Total Marks', 0)
                percentage = scores.get('Percentage', 0)
                grade = scores.get('Grade', 'N/A')
                
                subject_table += f"""
                <tr>
                    <td style="padding: 8px; border-bottom: 1px solid #ddd;"><strong>{subject_name}:</strong></td>
                    <td style="padding: 8px; border-bottom: 1px solid #ddd;">{total:.2f}/100</td>
                    <td style="padding: 8px; border-bottom: 1px solid #ddd;">{percentage:.2f}%</td>
                    <td style="padding: 8px; border-bottom: 1px solid #ddd; font-weight: bold; color: {'#27ae60' if grade in ['A+', 'A', 'B+'] else '#e74c3c'};">{grade}</td>
                </tr>
                """
            
            # Calculate overall performance if multiple subjects
            overall_percentage = 0
            if subject_results:
                overall_percentage = sum([s.get('Percentage', 0) for s in subject_results.values()]) / len(subject_results)
            
            overall_grade = self.calculate_overall_grade(overall_percentage)
            
            # Create email body
            body = f"""
<div style="font-family: Arial, sans-serif; max-width: 700px; margin: 0 auto;">
    <h2 style="color: #2c3e50;">🎓 Multi-Subject Exam Results</h2>
    <p>Dear <strong>{student_name}</strong>,</p>
    
    <p>Your exam evaluation for multiple subjects is complete. Here are your results:</p>
    
    <div style="background-color: #f8f9fa; padding: 20px; border-radius: 10px; margin: 20px 0;">
        <h3 style="color: #3498db; margin-top: 0;">📊 Subject-wise Results</h3>
        <table style="width: 100%; border-collapse: collapse;">
            <tr style="background-color: #e9ecef;">
                <th style="padding: 10px; text-align: left;">Subject</th>
                <th style="padding: 10px; text-align: left;">Total Marks</th>
                <th style="padding: 10px; text-align: left;">Percentage</th>
                <th style="padding: 10px; text-align: left;">Grade</th>
            </tr>
            {subject_table}
        </table>
        
        <div style="margin-top: 20px; padding: 15px; background-color: #d4edda; border-radius: 5px;">
            <h4 style="color: #155724; margin-top: 0;">🏆 Overall Performance</h4>
            <p><strong>Roll Number:</strong> {roll_no}</p>
            <p><strong>Overall Percentage:</strong> {overall_percentage:.2f}%</p>
            <p><strong>Overall Grade:</strong> <span style="font-weight: bold; color: {'#27ae60' if overall_grade in ['A+', 'A', 'B+'] else '#e74c3c'};">{overall_grade}</span></p>
        </div>
    </div>
    
    <div style="background-color: #e8f4fc; padding: 15px; border-radius: 8px; margin: 20px 0;">
        <h4 style="color: #2980b9; margin-top: 0;">📈 Performance Feedback</h4>
        <p>{self.get_performance_feedback(overall_percentage)}</p>
    </div>
    
    <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee;">
        <p><strong>Note:</strong> For detailed question-wise scores, please check the attached Excel file.</p>
        <p><em>This is an automated message. Please do not reply to this email.</em></p>
    </div>
    
    <div style="margin-top: 20px; color: #7f8c8d; font-size: 12px;">
        <p>Generated on: {datetime.now().strftime("%B %d, %Y at %I:%M %p")}</p>
        <p>© Multi-Subject Exam Evaluation System</p>
    </div>
</div>
"""
            
            message.attach(MIMEText(body, "html"))
            
            # Attach results file if available
            if results_file and os.path.exists(results_file):
                try:
                    with open(results_file, "rb") as file:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(file.read())
                        encoders.encode_base64(part)
                        part.add_header(
                            "Content-Disposition",
                            f"attachment; filename=Results_{student_name.replace(' ', '_')}.xlsx",
                        )
                        message.attach(part)
                except Exception as e:
                    return False, f"Could not attach file: {e}"
            
            # Send email
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(self.sender_email, self.app_password)
                server.sendmail(self.sender_email, student_email, message.as_string())
            
            self.sent_emails_log.append({
                'student': student_name,
                'email': student_email,
                'subjects': len(subject_results),
                'overall_percentage': overall_percentage,
                'time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'status': 'Sent'
            })
            
            return True, f"Email sent to {student_name}"
            
        except Exception as e:
            error_msg = str(e)
            self.sent_emails_log.append({
                'student': student_name,
                'email': student_email,
                'time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'status': 'Failed',
                'error': error_msg
            })
            return False, f"Failed to send email: {error_msg}"
    
    def calculate_overall_grade(self, percentage):
        """Calculate overall grade from percentage"""
        if percentage >= 90: return 'A+'
        elif percentage >= 80: return 'A'
        elif percentage >= 70: return 'B+'
        elif percentage >= 60: return 'B'
        elif percentage >= 50: return 'C'
        elif percentage >= 40: return 'D'
        else: return 'F'
    
    def get_performance_feedback(self, percentage):
        """Generate performance feedback"""
        if percentage >= 90:
            return "🌟 Outstanding performance across all subjects! Excellent understanding."
        elif percentage >= 80:
            return "🎯 Excellent overall performance! You have a strong grasp of the material."
        elif percentage >= 70:
            return "👍 Very good performance. Consistent results across subjects."
        elif percentage >= 60:
            return "📚 Good overall effort. Some subjects could use more attention."
        elif percentage >= 50:
            return "⚠️ Average overall performance. Consider focusing on weaker subjects."
        elif percentage >= 40:
            return "📉 Below average. Please focus on core concepts in all subjects."
        else:
            return "❌ Needs significant improvement. Consider seeking additional help in multiple subjects."

class MultiSubjectAnswerSheetEvaluator:
    """Main evaluation logic for multiple subjects"""
    
    def __init__(self, use_ocr=False):
        self.use_ocr = use_ocr
        self.subjects = []  # List of SubjectData objects
        self.consolidated_results_file = None
        self.log_messages = []
        self.email_sender = EmailSender(SENDER_EMAIL, APP_PASSWORD)
        self.pdf_processor = MultiSubjectPDFProcessor(OCR_API_KEY)
    
    def add_subject(self, name, master_pdf_path, student_pdfs=None):
        """Add a subject to evaluate"""
        subject = SubjectData(name, master_pdf_path, student_pdfs)
        self.subjects.append(subject)
        return subject
    
    def remove_subject(self, name):
        """Remove a subject by name"""
        self.subjects = [s for s in self.subjects if s.name != name]
    
    def get_subject(self, name):
        """Get subject by name"""
        for subject in self.subjects:
            if subject.name == name:
                return subject
        return None
    
    def load_subjects_from_csv(self, csv_path):
        """Load multiple subjects from CSV file"""
        try:
            df = pd.read_csv(csv_path)
            subjects = []
            
            for _, row in df.iterrows():
                subject_name = row['Subject_Name']
                master_pdf = row['Master_PDF']
                student_folder = row.get('Student_Folder', '')
                
                if os.path.exists(master_pdf) and os.path.exists(student_folder):
                    student_pdfs = glob.glob(os.path.join(student_folder, "*.pdf")) + \
                                   glob.glob(os.path.join(student_folder, "*.PDF"))
                    
                    subject = SubjectData(
                        name=subject_name,
                        master_pdf_path=master_pdf,
                        student_pdfs=student_pdfs
                    )
                    subjects.append(subject)
            
            self.subjects = subjects
            return True, f"Loaded {len(subjects)} subjects from CSV"
        except Exception as e:
            return False, f"Failed to load subjects: {str(e)}"
    
    def save_subjects_to_csv(self, csv_path):
        """Save subjects configuration to CSV"""
        try:
            data = []
            for subject in self.subjects:
                student_folder = os.path.dirname(subject.student_pdfs[0]) if subject.student_pdfs else ""
                data.append({
                    'Subject_Name': subject.name,
                    'Master_PDF': subject.master_pdf_path,
                    'Student_Folder': student_folder
                })
            
            df = pd.DataFrame(data)
            df.to_csv(csv_path, index=False)
            return True, f"Saved {len(self.subjects)} subjects to CSV"
        except Exception as e:
            return False, f"Failed to save subjects: {str(e)}"
    
    def log(self, message, widget=None):
        """Log message"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        print(log_entry)
        self.log_messages.append(log_entry)
        
        if widget:
            widget.insert(tk.END, log_entry + "\n")
            widget.see(tk.END)
    
    def extract_text_from_pdf(self, pdf_path, source_name, log_widget=None):
        """Extract text from PDF using either PyPDF2 or OCR"""
        # First try PyPDF2
        text = self.pdf_processor.extract_pdf_text(pdf_path)
        if text and len(text.strip()) > 100:
            return self.clean_extracted_text(text)
        
        # If PyPDF2 fails or text is too short, use OCR
        if self.use_ocr and OCR_API_KEY:
            self.log(f"  🔄 Using OCR for {source_name}", log_widget)
            ocr_text = self.pdf_processor.extract_text_with_ocr(pdf_path, log_widget)
            if ocr_text:
                return self.clean_extracted_text(ocr_text)
        
        return text
    
    def clean_extracted_text(self, text):
        """Clean extracted text"""
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'(Q\d+)', r'\n\1:', text)
        return text.strip()
    
    def parse_master_answer(self, text):
        """Parse master answers from text"""
        master_answers = {}
        pattern = r'Q(\d+)([a-zA-Z])?:\s*(.*?)(?=Q\d+[a-zA-Z]?:|Q\d+:|$)'
        matches = re.findall(pattern, text, re.DOTALL | re.IGNORECASE)
        
        for q_num, subpart, answer in matches:
            q_num = int(q_num)
            answer = answer.strip()
            
            if subpart:
                key = f"{q_num}{subpart.lower()}"
            else:
                key = str(q_num)
            
            master_answers[key] = answer
        
        return master_answers
    
    def extract_student_info(self, text):
        """Extract student information from text"""
        info = {'name': '', 'roll_no': '', 'email': ''}
        
        # Extract name
        name_patterns = [
            r'Student[^:\n]*:\s*([^\n(]+)',
            r'Name[^:\n]*:\s*([^\n(]+)',
            r'Student\s*name[^:\n]*:\s*([^\n(]+)'
        ]
        
        for pattern in name_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                info['name'] = match.group(1).strip()
                info['name'] = re.sub(r'\s*\(.*\)', '', info['name'])
                break
        
        # Extract roll number
        roll_patterns = [
            r'Roll\s*(?:No|no|No\.?|Number)[:\s]*([A-Za-z0-9-]+)',
            r'Roll[^:\n]*:\s*([A-Za-z0-9-]+)'
        ]
        
        for pattern in roll_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                info['roll_no'] = match.group(1).strip()
                break
        
        # Extract email
        email_patterns = [
            r'Email[^:\n]*:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})',
            r'Email\s*id[^:\n]*:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})',
            r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
        ]
        
        for pattern in email_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                info['email'] = match.group(1).strip().lower()
                break
        
        return info
    
    def parse_student_answers(self, text, student_id):
        """Parse student answers and info"""
        info = self.extract_student_info(text)
        
        # Set defaults if not found
        if not info['name']:
            info['name'] = f'Student_{student_id}'
        if not info['roll_no']:
            info['roll_no'] = f'ID_{student_id}'
        if not info['email']:
            info['email'] = f'student_{student_id}@example.com'
        
        # Extract answers
        answers = {}
        pattern = r'Q(\d+)([a-zA-Z])?:\s*(.*?)(?=Q\d+[a-zA-Z]?:|Q\d+:|$)'
        matches = re.findall(pattern, text, re.DOTALL | re.IGNORECASE)
        
        for q_num, subpart, answer in matches:
            q_num = int(q_num)
            answer = answer.strip()
            
            if subpart:
                key = f"{q_num}{subpart.lower()}"
            else:
                key = str(q_num)
            
            if answer and len(answer) > 3:
                answers[key] = answer
        
        return {
            'name': info['name'],
            'roll_no': info['roll_no'],
            'email': info['email'],
            'answers': answers
        }
    
    def evaluate_answer(self, master_answer, student_answer):
        """Evaluate a single answer using AI-based scoring"""
        if not student_answer or student_answer.strip() == "":
            return 0.0
        
        try:
            def preprocess(text):
                text = text.lower()
                text = re.sub(r'[^\w\s]', ' ', text)
                text = re.sub(r'\s+', ' ', text).strip()
                return text
            
            master_processed = preprocess(master_answer)
            student_processed = preprocess(student_answer)
            
            if not master_processed or not student_processed:
                return 0.0
            
            # Calculate word overlap
            master_words = set(master_processed.split()[:20])
            student_words = set(student_processed.split()[:20])
            
            if not master_words:
                return 0.0
            
            overlap = len(master_words.intersection(student_words)) / len(master_words)
            base_score = overlap * 10
            
            # Apply adjustments
            final_score = base_score
            
            # Length-based adjustments
            if len(student_answer.strip()) > 30:
                final_score = min(final_score + 1.5, 10)
            elif len(student_answer.strip()) > 15:
                final_score = min(final_score + 0.5, 10)
            
            if len(student_answer.strip()) > 5:
                final_score = max(final_score, 2.0)
            
            return round(final_score, 2)
        except:
            return 5.0
    
    def calculate_grade(self, percentage):
        """Calculate grade from percentage"""
        if percentage >= 90: return 'A+'
        elif percentage >= 80: return 'A'
        elif percentage >= 70: return 'B+'
        elif percentage >= 60: return 'B'
        elif percentage >= 50: return 'C'
        elif percentage >= 40: return 'D'
        else: return 'F'
    
    def evaluate_subject(self, subject_data, log_widget=None):
        """Evaluate a single subject"""
        self.log(f"\n📚 EVALUATING SUBJECT: {subject_data.name}", log_widget)
        self.log(f"Master PDF: {os.path.basename(subject_data.master_pdf_path)}", log_widget)
        self.log(f"Student PDFs: {len(subject_data.student_pdfs)} files", log_widget)
        
        # Extract Master Answers
        master_text = self.extract_text_from_pdf(subject_data.master_pdf_path, 
                                                f"Master - {subject_data.name}", 
                                                log_widget)
        if not master_text:
            raise Exception(f"Failed to extract text from master PDF for {subject_data.name}")
        
        master_answers = self.parse_master_answer(master_text)
        if not master_answers:
            raise Exception(f"Could not parse questions from master for {subject_data.name}")
        
        self.log(f"✓ Found {len(master_answers)} questions in master", log_widget)
        
        # Process Student Answer Sheets
        results = []
        max_score = 10 * len(master_answers)
        
        for i, student_pdf in enumerate(subject_data.student_pdfs):
            filename = os.path.basename(student_pdf)
            self.log(f"\n🔍 Processing {filename} ({i+1}/{len(subject_data.student_pdfs)})", log_widget)
            
            student_text = self.extract_text_from_pdf(student_pdf, filename, log_widget)
            if not student_text:
                self.log(f"  ✗ Failed to extract text", log_widget)
                continue
            
            student_data = self.parse_student_answers(student_text, i+1)
            
            # Evaluate answers
            total_score = 0
            question_scores = {}
            
            for q_key, m_ans in master_answers.items():
                s_ans = student_data['answers'].get(q_key, '')
                score = self.evaluate_answer(m_ans, s_ans)
                question_scores[f"Q{q_key}"] = score
                total_score += score
            
            percentage = round((total_score / max_score) * 100, 2) if max_score > 0 else 0
            grade = self.calculate_grade(percentage)
            
            result = {
                'Subject': subject_data.name,
                'Name': student_data['name'],
                'Roll No': student_data['roll_no'],
                'Email': student_data['email'],
                'Total Marks': round(total_score, 2),
                'Percentage': percentage,
                'Grade': grade,
                **question_scores
            }
            results.append(result)
            
            self.log(f"  👤 Name: {student_data['name']}", log_widget)
            self.log(f"  📝 Roll No: {student_data['roll_no']}", log_widget)
            self.log(f"  📊 Score: {total_score:.1f}/{max_score} ({percentage}%)", log_widget)
            self.log(f"  🏆 Grade: {grade}", log_widget)
        
        subject_data.results = results
        return results
    
    def evaluate_all_subjects(self, log_widget=None, progress_callback=None):
        """Evaluate all subjects"""
        self.log("=" * 60, log_widget)
        self.log("📚 STARTING MULTI-SUBJECT EVALUATION", log_widget)
        self.log(f"Subjects: {len(self.subjects)}", log_widget)
        self.log("=" * 60, log_widget)
        
        # Create output directory
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
        
        all_results = []
        
        for idx, subject in enumerate(self.subjects):
            if progress_callback:
                progress_callback((idx + 1) / len(self.subjects) * 50)
            
            try:
                subject_results = self.evaluate_subject(subject, log_widget)
                all_results.extend(subject_results)
                
                # Save individual subject results
                subject.results_file = self.save_subject_results(subject, log_widget)
                
            except Exception as e:
                self.log(f"❌ Error evaluating {subject.name}: {str(e)}", log_widget)
                continue
        
        # Step 3: Save Consolidated Results
        self.log("\nSTEP 3: SAVING CONSOLIDATED RESULTS", log_widget)
        
        if all_results:
            self.consolidated_results_file = self.save_consolidated_results(all_results, log_widget)
            
            # Generate summary statistics
            self.generate_multi_subject_summary(all_results, log_widget)
            
            if progress_callback:
                progress_callback(100)
        else:
            raise Exception("No results to evaluate")
        
        return all_results
    
    def save_subject_results(self, subject_data, log_widget=None):
        """Save individual subject results to Excel"""
        if not subject_data.results:
            return None
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(OUTPUT_DIR, f"{subject_data.name}_results_{timestamp}.xlsx")
        
        df = pd.DataFrame(subject_data.results)
        
        # Save to Excel
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Results', index=False)
            
            worksheet = writer.sheets['Results']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 30)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        self.log(f"✓ {subject_data.name} results saved to: {filename}", log_widget)
        return filename
    
    def save_consolidated_results(self, all_results, log_widget=None):
        """Save consolidated results from all subjects"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(OUTPUT_DIR, f"consolidated_results_{timestamp}.xlsx")
        
        df = pd.DataFrame(all_results)
        
        # Save to Excel with multiple sheets
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            # Consolidated sheet
            df.to_excel(writer, sheet_name='All Results', index=False)
            
            # Individual subject sheets
            for subject in self.subjects:
                if subject.results:
                    subject_df = pd.DataFrame(subject.results)
                    subject_df.to_excel(writer, sheet_name=subject.name[:31], index=False)
            
            # Summary sheet
            summary_data = self.generate_summary_dataframe(all_results)
            summary_data.to_excel(writer, sheet_name='Summary', index=False)
        
        self.log(f"✓ Consolidated results saved to: {filename}", log_widget)
        return filename
    
    def generate_summary_dataframe(self, all_results):
        """Generate summary dataframe"""
        if not all_results:
            return pd.DataFrame()
        
        # Group by student
        student_data = {}
        for result in all_results:
            roll_no = result['Roll No']
            name = result['Name']
            email = result['Email']
            subject = result['Subject']
            total = result['Total Marks']
            percentage = result['Percentage']
            grade = result['Grade']
            
            if roll_no not in student_data:
                student_data[roll_no] = {
                    'Roll No': roll_no,
                    'Name': name,
                    'Email': email,
                    'Subjects': []
                }
            
            student_data[roll_no]['Subjects'].append({
                'name': subject,
                'total': total,
                'percentage': percentage,
                'grade': grade
            })
        
        # Create summary rows
        summary_rows = []
        for roll_no, data in student_data.items():
            row = {
                'Roll No': roll_no,
                'Name': data['Name'],
                'Email': data['Email'],
                'Total Subjects': len(data['Subjects'])
            }
            
            # Add subject-wise columns
            for subj_data in data['Subjects']:
                subject_name = subj_data['name']
                row[f'{subject_name}_Marks'] = subj_data['total']
                row[f'{subject_name}_%'] = subj_data['percentage']
                row[f'{subject_name}_Grade'] = subj_data['grade']
            
            # Calculate overall
            if data['Subjects']:
                overall_percentage = sum([s['percentage'] for s in data['Subjects']]) / len(data['Subjects'])
                row['Overall_%'] = overall_percentage
                row['Overall_Grade'] = self.calculate_grade(overall_percentage)
            
            summary_rows.append(row)
        
        return pd.DataFrame(summary_rows)
    
    def send_emails(self, all_results, log_widget=None, progress_callback=None):
        """Send emails to all students with multi-subject results"""
        self.log("\n" + "=" * 60, log_widget)
        self.log("📧 SENDING EMAILS TO STUDENTS (MULTI-SUBJECT)", log_widget)
        self.log("=" * 60, log_widget)
        
        # Test email connection
        self.log("Testing email connection...", log_widget)
        success, message = self.email_sender.test_connection()
        if not success:
            self.log(f"✗ {message}", log_widget)
            return 0, len(set([r['Roll No'] for r in all_results]))
        
        self.log(f"✓ {message}", log_widget)
        
        # Group results by student
        student_results = {}
        for result in all_results:
            roll_no = result['Roll No']
            if roll_no not in student_results:
                student_results[roll_no] = {
                    'Name': result['Name'],
                    'Email': result['Email'],
                    'Roll No': roll_no,
                    'subjects': {}
                }
            
            student_results[roll_no]['subjects'][result['Subject']] = {
                'Total Marks': result['Total Marks'],
                'Percentage': result['Percentage'],
                'Grade': result['Grade']
            }
        
        success_count = 0
        fail_count = 0
        
        for i, (roll_no, student_data) in enumerate(student_results.items()):
            self.log(f"\n📨 Sending to {student_data['Name']} ({i+1}/{len(student_results)})", log_widget)
            self.log(f"  Email: {student_data['Email']}", log_widget)
            self.log(f"  Subjects: {', '.join(student_data['subjects'].keys())}", log_widget)
            
            if progress_callback:
                progress_callback((i + 1) / len(student_results) * 100)
            
            success, msg = self.email_sender.send_results_email(
                student_data, 
                student_data['subjects'], 
                self.consolidated_results_file
            )
            
            if success:
                self.log(f"  ✅ {msg}", log_widget)
                success_count += 1
            else:
                self.log(f"  ❌ {msg}", log_widget)
                fail_count += 1
        
        # Save email log
        self.save_email_log(log_widget)
        
        return success_count, fail_count
    
    def save_email_log(self, log_widget=None):
        """Save email sending log"""
        if self.email_sender.sent_emails_log:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_file = os.path.join(OUTPUT_DIR, f"email_log_{timestamp}.csv")
            
            df = pd.DataFrame(self.email_sender.sent_emails_log)
            df.to_csv(log_file, index=False)
            self.log(f"\n📝 Email log saved to: {log_file}", log_widget)
    
    def generate_multi_subject_summary(self, all_results, log_widget=None):
        """Display summary statistics for multiple subjects"""
        self.log("\n" + "=" * 60, log_widget)
        self.log("📊 MULTI-SUBJECT EVALUATION SUMMARY", log_widget)
        self.log("=" * 60, log_widget)
        
        if not all_results:
            self.log("No results to display", log_widget)
            return
        
        # Subject-wise summary
        self.log("\nSUBJECT-WISE SUMMARY:", log_widget)
        for subject in self.subjects:
            if subject.results:
                subject_name = subject.name
                percentages = [r['Percentage'] for r in subject.results]
                
                self.log(f"\n{subject_name}:", log_widget)
                self.log(f"  Students: {len(percentages)}", log_widget)
                self.log(f"  Average Score: {sum(percentages)/len(percentages):.1f}%", log_widget)
                self.log(f"  Highest Score: {max(percentages):.1f}%", log_widget)
                self.log(f"  Lowest Score: {min(percentages):.1f}%", log_widget)
                
                # Grade distribution
                grade_dist = {}
                for result in subject.results:
                    grade = result['Grade']
                    grade_dist[grade] = grade_dist.get(grade, 0) + 1
                
                for grade, count in sorted(grade_dist.items()):
                    self.log(f"    {grade}: {count} student(s)", log_widget)
        
        # Overall summary
        self.log("\nOVERALL SUMMARY:", log_widget)
        unique_students = len(set([r['Roll No'] for r in all_results]))
        self.log(f"Total Unique Students: {unique_students}", log_widget)
        self.log(f"Total Subjects Evaluated: {len(self.subjects)}", log_widget)
        self.log(f"Total Answer Sheets Processed: {len(all_results)}", log_widget)

class MultiSubjectGUI:
    """GUI Application for Multi-Subject AI-based Answer Sheet Evaluation"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-Subject AI Answer Sheet Evaluation System")
        self.root.geometry("1200x800")
        
        # Variables
        self.subjects = []  # List of SubjectData objects
        self.use_ocr = tk.BooleanVar(value=True)
        self.send_emails = tk.BooleanVar(value=False)
        self.current_subject_index = -1
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_subject_manager_tab()
        self.create_evaluation_tab()
        self.create_pdf_processing_tab()
        self.create_settings_tab()
        
        # Status bar
        self.status_bar = tk.Label(root, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Initialize evaluator
        self.evaluator = MultiSubjectAnswerSheetEvaluator(use_ocr=self.use_ocr.get())
    
    def create_subject_manager_tab(self):
        """Create the subject management tab"""
        self.subject_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.subject_frame, text="Subject Manager")
        
        # Title
        title_label = tk.Label(self.subject_frame, text="📚 Multi-Subject Management", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Main content frame
        content_frame = ttk.Frame(self.subject_frame)
        content_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Left panel - Subject List
        left_panel = ttk.LabelFrame(content_frame, text="Subjects List", padding=10)
        left_panel.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # Subject listbox with scrollbar
        listbox_frame = ttk.Frame(left_panel)
        listbox_frame.pack(fill='both', expand=True, pady=5)
        
        self.subject_listbox = tk.Listbox(listbox_frame, height=15, font=("Arial", 10))
        self.subject_listbox.pack(side='left', fill='both', expand=True)
        
        listbox_scrollbar = ttk.Scrollbar(listbox_frame, orient='vertical')
        listbox_scrollbar.pack(side='right', fill='y')
        self.subject_listbox.config(yscrollcommand=listbox_scrollbar.set)
        listbox_scrollbar.config(command=self.subject_listbox.yview)
        
        # Subject list buttons
        list_buttons_frame = ttk.Frame(left_panel)
        list_buttons_frame.pack(fill='x', pady=5)
        
        ttk.Button(list_buttons_frame, text="🔄 Refresh List", 
                  command=self.refresh_subject_list).pack(side='left', padx=2)
        ttk.Button(list_buttons_frame, text="🗑️ Remove Selected", 
                  command=self.remove_selected_subject).pack(side='left', padx=2)
        ttk.Button(list_buttons_frame, text="📋 Clear All", 
                  command=self.clear_all_subjects).pack(side='left', padx=2)
        
        # Right panel - Add/Edit Subject
        right_panel = ttk.LabelFrame(content_frame, text="Add/Edit Subject", padding=10)
        right_panel.pack(side='right', fill='both', expand=True)
        
        # Subject name
        ttk.Label(right_panel, text="Subject Name:").grid(row=0, column=0, sticky='w', pady=5)
        self.subject_name_var = tk.StringVar()
        ttk.Entry(right_panel, textvariable=self.subject_name_var, width=30).grid(row=0, column=1, padx=5, pady=5)
        
        # Master PDF
        ttk.Label(right_panel, text="Master Answer Sheet:").grid(row=1, column=0, sticky='w', pady=5)
        self.master_pdf_var = tk.StringVar()
        master_frame = ttk.Frame(right_panel)
        master_frame.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        ttk.Entry(master_frame, textvariable=self.master_pdf_var, width=25).pack(side='left', fill='x', expand=True)
        ttk.Button(master_frame, text="Browse", 
                  command=self.browse_master_pdf).pack(side='left', padx=5)
        
        # Student PDFs folder
        ttk.Label(right_panel, text="Student Answer Sheets Folder:").grid(row=2, column=0, sticky='w', pady=5)
        self.student_folder_var = tk.StringVar()
        student_frame = ttk.Frame(right_panel)
        student_frame.grid(row=2, column=1, padx=5, pady=5, sticky='ew')
        ttk.Entry(student_frame, textvariable=self.student_folder_var, width=25).pack(side='left', fill='x', expand=True)
        ttk.Button(student_frame, text="Browse", 
                  command=self.browse_student_folder).pack(side='left', padx=5)
        
        # Control buttons
        buttons_frame = ttk.Frame(right_panel)
        buttons_frame.grid(row=3, column=0, columnspan=2, pady=15)
        
        ttk.Button(buttons_frame, text="➕ Add Subject", 
                  command=self.add_subject, style='Accent.TButton').pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="✏️ Update Subject", 
                  command=self.update_subject).pack(side='left', padx=5)
        
        # Import/Export buttons
        import_export_frame = ttk.Frame(right_panel)
        import_export_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        ttk.Button(import_export_frame, text="📥 Import from CSV", 
                  command=self.import_subjects_csv).pack(side='left', padx=5)
        ttk.Button(import_export_frame, text="📤 Export to CSV", 
                  command=self.export_subjects_csv).pack(side='left', padx=5)
        
        # Subject count
        self.subject_count_label = tk.Label(right_panel, text="Subjects: 0", font=("Arial", 10))
        self.subject_count_label.grid(row=5, column=0, columnspan=2, pady=10)
        
        # Bind listbox selection
        self.subject_listbox.bind('<<ListboxSelect>>', self.on_subject_select)
    
    def create_evaluation_tab(self):
        """Create the evaluation tab"""
        self.eval_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.eval_frame, text="Evaluation")
        
        # Title
        title_label = tk.Label(self.eval_frame, text="🎓 Multi-Subject Evaluation", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Main content frame
        content_frame = ttk.Frame(self.eval_frame)
        content_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Left panel for controls
        left_panel = ttk.Frame(content_frame)
        left_panel.pack(side='left', fill='y', padx=(0, 10))
        
        # Evaluation summary
        summary_frame = ttk.LabelFrame(left_panel, text="Evaluation Summary", padding=10)
        summary_frame.pack(fill='x', pady=(0, 10))
        
        self.summary_text = tk.Text(summary_frame, height=6, width=40)
        self.summary_text.pack(fill='both', expand=True)
        self.summary_text.config(state='disabled')
        
        # Options frame
        options_frame = ttk.LabelFrame(left_panel, text="Evaluation Options", padding=10)
        options_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Checkbutton(options_frame, text="Use OCR (for handwritten sheets)", 
                       variable=self.use_ocr).pack(anchor='w', pady=2)
        ttk.Checkbutton(options_frame, text="Send Emails to Students", 
                       variable=self.send_emails).pack(anchor='w', pady=2)
        
        # Control buttons
        control_frame = ttk.Frame(left_panel)
        control_frame.pack(fill='x', pady=5)
        
        ttk.Button(control_frame, text="📊 Evaluate All Subjects", 
                  command=self.start_evaluation, style='Accent.TButton').pack(fill='x', pady=5)
        ttk.Button(control_frame, text="📧 Send Emails Only", 
                  command=self.send_emails_only).pack(fill='x', pady=2)
        ttk.Button(control_frame, text="📁 Open Results Folder", 
                  command=self.open_results_folder).pack(fill='x', pady=2)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(left_panel, mode='determinate')
        self.progress_bar.pack(fill='x', pady=10)
        
        # Right panel for logs
        right_panel = ttk.Frame(content_frame)
        right_panel.pack(side='right', fill='both', expand=True)
        
        ttk.Label(right_panel, text="Evaluation Log", font=("Arial", 11, "bold")).pack(anchor='w', pady=(0, 5))
        
        # Log text area with scrollbar
        log_frame = ttk.Frame(right_panel)
        log_frame.pack(fill='both', expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=25, width=70)
        self.log_text.pack(fill='both', expand=True)
        
        # Clear log button
        ttk.Button(right_panel, text="Clear Log", command=self.clear_log).pack(anchor='e', pady=5)
    
    def create_pdf_processing_tab(self):
        """Create PDF processing tab"""
        self.pdf_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.pdf_frame, text="PDF Processing")
        
        title_label = tk.Label(self.pdf_frame, text="📄 PDF OCR Processing", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        content_frame = ttk.Frame(self.pdf_frame)
        content_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Input section
        input_frame = ttk.LabelFrame(content_frame, text="Input", padding=10)
        input_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(input_frame, text="Select PDF file or folder:").pack(anchor='w')
        
        self.pdf_input_path = tk.StringVar()
        input_path_frame = ttk.Frame(input_frame)
        input_path_frame.pack(fill='x', pady=5)
        
        ttk.Entry(input_path_frame, textvariable=self.pdf_input_path, width=40).pack(side='left', fill='x', expand=True)
        ttk.Button(input_path_frame, text="Browse File", command=self.browse_pdf_file).pack(side='left', padx=2)
        ttk.Button(input_path_frame, text="Browse Folder", command=self.browse_pdf_folder).pack(side='left', padx=2)
        
        # Output options
        options_frame = ttk.LabelFrame(content_frame, text="Output Options", padding=10)
        options_frame.pack(fill='x', pady=(0, 10))
        
        self.create_searchable_pdf = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Create searchable PDF", 
                       variable=self.create_searchable_pdf).pack(anchor='w', pady=2)
        
        # Process button
        ttk.Button(content_frame, text="🔄 Process PDFs with OCR", 
                  command=self.process_pdfs_ocr, style='Accent.TButton').pack(fill='x', pady=10)
        
        # PDF processing log
        ttk.Label(content_frame, text="Processing Log", font=("Arial", 11, "bold")).pack(anchor='w', pady=(10, 5))
        
        self.pdf_log_text = scrolledtext.ScrolledText(content_frame, height=15, width=80)
        self.pdf_log_text.pack(fill='both', expand=True)
        
        ttk.Button(content_frame, text="Clear Log", command=lambda: self.pdf_log_text.delete(1.0, tk.END)).pack(anchor='e', pady=5)
    
    def create_settings_tab(self):
        """Create settings tab"""
        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text="Settings")
        
        title_label = tk.Label(self.settings_frame, text="⚙️ System Settings", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        content_frame = ttk.Frame(self.settings_frame)
        content_frame.pack(fill='both', expand=True, padx=50, pady=20)
        
        # Email settings
        email_frame = ttk.LabelFrame(content_frame, text="Email Settings", padding=15)
        email_frame.pack(fill='x', pady=(0, 15))
        
        ttk.Label(email_frame, text="Sender Email:").grid(row=0, column=0, sticky='w', pady=5)
        self.sender_email_var = tk.StringVar(value=SENDER_EMAIL)
        ttk.Entry(email_frame, textvariable=self.sender_email_var, width=40).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(email_frame, text="App Password:").grid(row=1, column=0, sticky='w', pady=5)
        self.app_password_var = tk.StringVar(value=APP_PASSWORD)
        ttk.Entry(email_frame, textvariable=self.app_password_var, show="*", width=40).grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Button(email_frame, text="Test Email Connection", 
                  command=self.test_email_connection).grid(row=2, column=0, columnspan=2, pady=10)
        
        # OCR settings
        ocr_frame = ttk.LabelFrame(content_frame, text="OCR Settings", padding=15)
        ocr_frame.pack(fill='x', pady=(0, 15))
        
        ttk.Label(ocr_frame, text="OCR API Key:").grid(row=0, column=0, sticky='w', pady=5)
        self.ocr_api_key_var = tk.StringVar(value=OCR_API_KEY)
        ttk.Entry(ocr_frame, textvariable=self.ocr_api_key_var, width=40).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(ocr_frame, text="Output Directory:").grid(row=1, column=0, sticky='w', pady=5)
        self.output_dir_var = tk.StringVar(value=OUTPUT_DIR)
        ttk.Entry(ocr_frame, textvariable=self.output_dir_var, width=40).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(ocr_frame, text="Browse", 
                  command=lambda: self.output_dir_var.set(filedialog.askdirectory())).grid(row=1, column=2, padx=5)
        
        # Save button
        ttk.Button(content_frame, text="💾 Save Settings", 
                  command=self.save_settings, style='Accent.TButton').pack(pady=20)
    
    def add_subject(self):
        """Add a new subject"""
        name = self.subject_name_var.get().strip()
        master_pdf = self.master_pdf_var.get().strip()
        student_folder = self.student_folder_var.get().strip()
        
        if not name:
            messagebox.showerror("Error", "Please enter a subject name")
            return
        
        if not master_pdf or not os.path.exists(master_pdf):
            messagebox.showerror("Error", "Please select a valid master PDF file")
            return
        
        if not student_folder or not os.path.exists(student_folder):
            messagebox.showerror("Error", "Please select a valid student folder")
            return
        
        # Get student PDFs from folder
        student_pdfs = glob.glob(os.path.join(student_folder, "*.pdf")) + \
                      glob.glob(os.path.join(student_folder, "*.PDF"))
        
        if not student_pdfs:
            messagebox.showwarning("Warning", "No PDF files found in the student folder")
            return
        
        # Check if subject already exists
        for subject in self.subjects:
            if subject.name == name:
                messagebox.showwarning("Warning", f"Subject '{name}' already exists")
                return
        
        # Add subject
        subject = SubjectData(name, master_pdf, student_pdfs)
        self.subjects.append(subject)
        
        # Update evaluator
        self.evaluator.add_subject(name, master_pdf, student_pdfs)
        
        # Update UI
        self.refresh_subject_list()
        self.update_summary()
        
        # Clear form
        self.subject_name_var.set("")
        self.master_pdf_var.set("")
        self.student_folder_var.set("")
        
        self.log_message(f"✅ Added subject: {name} with {len(student_pdfs)} student PDFs")
    
    def update_subject(self):
        """Update selected subject"""
        if self.current_subject_index < 0:
            messagebox.showwarning("Warning", "Please select a subject to update")
            return
        
        name = self.subject_name_var.get().strip()
        master_pdf = self.master_pdf_var.get().strip()
        student_folder = self.student_folder_var.get().strip()
        
        if not all([name, master_pdf, student_folder]):
            messagebox.showerror("Error", "Please fill all fields")
            return
        
        # Get student PDFs
        student_pdfs = glob.glob(os.path.join(student_folder, "*.pdf")) + \
                      glob.glob(os.path.join(student_folder, "*.PDF"))
        
        # Update subject
        self.subjects[self.current_subject_index].name = name
        self.subjects[self.current_subject_index].master_pdf_path = master_pdf
        self.subjects[self.current_subject_index].student_pdfs = student_pdfs
        
        # Update evaluator
        self.evaluator.remove_subject(self.subjects[self.current_subject_index].name)
        self.evaluator.add_subject(name, master_pdf, student_pdfs)
        
        # Update UI
        self.refresh_subject_list()
        self.update_summary()
        
        self.log_message(f"✅ Updated subject: {name}")
    
    def remove_selected_subject(self):
        """Remove selected subject"""
        selection = self.subject_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a subject to remove")
            return
        
        index = selection[0]
        subject_name = self.subjects[index].name
        
        # Remove from list
        self.subjects.pop(index)
        self.evaluator.remove_subject(subject_name)
        
        # Update UI
        self.refresh_subject_list()
        self.update_summary()
        self.clear_subject_form()
        
        self.log_message(f"✅ Removed subject: {subject_name}")
    
    def clear_all_subjects(self):
        """Clear all subjects"""
        if not self.subjects:
            return
        
        if messagebox.askyesno("Confirm", "Are you sure you want to remove all subjects?"):
            self.subjects.clear()
            self.evaluator.subjects.clear()
            self.refresh_subject_list()
            self.update_summary()
            self.clear_subject_form()
            self.log_message("✅ Cleared all subjects")
    
    def refresh_subject_list(self):
        """Refresh the subject listbox"""
        self.subject_listbox.delete(0, tk.END)
        for subject in self.subjects:
            student_count = len(subject.student_pdfs)
            self.subject_listbox.insert(tk.END, f"{subject.name} ({student_count} students)")
        
        self.subject_count_label.config(text=f"Subjects: {len(self.subjects)}")
    
    def on_subject_select(self, event):
        """Handle subject selection"""
        selection = self.subject_listbox.curselection()
        if selection:
            index = selection[0]
            self.current_subject_index = index
            subject = self.subjects[index]
            
            # Populate form
            self.subject_name_var.set(subject.name)
            self.master_pdf_var.set(subject.master_pdf_path)
            
            # Get folder from first student PDF
            if subject.student_pdfs:
                folder = os.path.dirname(subject.student_pdfs[0])
                self.student_folder_var.set(folder)
            else:
                self.student_folder_var.set("")
    
    def clear_subject_form(self):
        """Clear the subject form"""
        self.subject_name_var.set("")
        self.master_pdf_var.set("")
        self.student_folder_var.set("")
        self.current_subject_index = -1
    
    def update_summary(self):
        """Update evaluation summary"""
        total_subjects = len(self.subjects)
        total_students = sum([len(s.student_pdfs) for s in self.subjects])
        
        summary_text = f"""Subjects: {total_subjects}
Total Answer Sheets: {total_students}

Subjects List:
"""
        for subject in self.subjects:
            summary_text += f"• {subject.name}: {len(subject.student_pdfs)} students\n"
        
        self.summary_text.config(state='normal')
        self.summary_text.delete(1.0, tk.END)
        self.summary_text.insert(1.0, summary_text)
        self.summary_text.config(state='disabled')
    
    def browse_master_pdf(self):
        """Browse for master PDF file"""
        file_path = filedialog.askopenfilename(
            title="Select Master Answer Sheet",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.master_pdf_var.set(file_path)
    
    def browse_student_folder(self):
        """Browse for student folder"""
        folder_path = filedialog.askdirectory(title="Select Folder with Student Answer Sheets")
        if folder_path:
            self.student_folder_var.set(folder_path)
    
    def import_subjects_csv(self):
        """Import subjects from CSV file"""
        file_path = filedialog.askopenfilename(
            title="Import Subjects from CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            success, message = self.evaluator.load_subjects_from_csv(file_path)
            if success:
                # Update local subjects list
                self.subjects = self.evaluator.subjects.copy()
                self.refresh_subject_list()
                self.update_summary()
                messagebox.showinfo("Success", message)
                self.log_message(f"✅ {message}")
            else:
                messagebox.showerror("Error", message)
                self.log_message(f"❌ {message}")
    
    def export_subjects_csv(self):
        """Export subjects to CSV file"""
        if not self.subjects:
            messagebox.showwarning("Warning", "No subjects to export")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Export Subjects to CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            success, message = self.evaluator.save_subjects_to_csv(file_path)
            if success:
                messagebox.showinfo("Success", message)
                self.log_message(f"✅ {message}")
            else:
                messagebox.showerror("Error", message)
                self.log_message(f"❌ {message}")
    
    def start_evaluation(self):
        """Start the evaluation process in a separate thread"""
        if not self.subjects:
            messagebox.showerror("Error", "Please add at least one subject")
            return
        
        # Validate all subjects
        valid_subjects = []
        for subject in self.subjects:
            if not os.path.exists(subject.master_pdf_path):
                self.log_message(f"Warning: Master PDF not found for {subject.name}")
                continue
            
            valid_pdfs = [p for p in subject.student_pdfs if os.path.exists(p)]
            if not valid_pdfs:
                self.log_message(f"Warning: No valid student PDFs for {subject.name}")
                continue
            
            subject.student_pdfs = valid_pdfs
            valid_subjects.append(subject)
        
        if not valid_subjects:
            messagebox.showerror("Error", "No valid subjects to evaluate")
            return
        
        # Disable controls during evaluation
        self.toggle_controls(False)
        self.progress_bar['value'] = 0
        self.log_message("=" * 60)
        self.log_message("🚀 Starting Multi-Subject Evaluation")
        self.log_message(f"Total Subjects: {len(valid_subjects)}")
        self.log_message(f"Total Answer Sheets: {sum([len(s.student_pdfs) for s in valid_subjects])}")
        self.log_message(f"OCR Enabled: {self.use_ocr.get()}")
        self.log_message("=" * 60)
        
        # Update evaluator
        self.evaluator.subjects = valid_subjects
        self.evaluator.use_ocr = self.use_ocr.get()
        
        # Start evaluation in separate thread
        thread = threading.Thread(target=self.run_evaluation)
        thread.daemon = True
        thread.start()
    
    def run_evaluation(self):
        """Run the evaluation process"""
        try:
            # Run evaluation with progress updates
            def update_progress(value):
                self.root.after(0, lambda: self.progress_bar.config(value=value))
            
            results = self.evaluator.evaluate_all_subjects(
                self.log_text,
                update_progress
            )
            
            # Send emails if requested
            if self.send_emails.get():
                success_count, fail_count = self.evaluator.send_emails(
                    results,
                    self.log_text,
                    update_progress
                )
                
                self.log_message(f"\n📧 Email Summary:")
                self.log_message(f"  ✅ Successfully sent: {success_count}")
                self.log_message(f"  ❌ Failed to send: {fail_count}")
            
            # Enable controls
            self.root.after(0, lambda: self.toggle_controls(True))
            
            # Show success message
            self.root.after(0, lambda: messagebox.showinfo(
                "Success",
                f"Evaluation completed successfully!\n\n"
                f"Results saved to: {self.evaluator.consolidated_results_file}"
            ))
            
            # Update status
            self.root.after(0, lambda: self.status_bar.config(text="Evaluation completed"))
            
        except Exception as e:
            error_msg = str(e)
            self.log_message(f"\n❌ Evaluation failed: {error_msg}")
            self.root.after(0, lambda: self.toggle_controls(True))
            self.root.after(0, lambda: messagebox.showerror("Error", f"Evaluation failed: {error_msg}"))
            self.root.after(0, lambda: self.status_bar.config(text="Evaluation failed"))
    
    def send_emails_only(self):
        """Send emails for existing results"""
        if not hasattr(self.evaluator, 'consolidated_results_file') or not self.evaluator.consolidated_results_file:
            messagebox.showinfo("Info", "Please run evaluation first to generate results.")
            return
        
        # Load results from file and send emails
        messagebox.showinfo("Info", "This feature would send emails based on existing results.")
    
    def browse_pdf_file(self):
        """Browse for PDF file for OCR processing"""
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.pdf_input_path.set(file_path)
            self.pdf_log_message(f"Selected PDF: {os.path.basename(file_path)}")
    
    def browse_pdf_folder(self):
        """Browse for folder containing PDFs for OCR processing"""
        folder_path = filedialog.askdirectory(title="Select Folder with PDFs")
        if folder_path:
            self.pdf_input_path.set(folder_path)
            self.pdf_log_message(f"Selected folder: {folder_path}")
    
    def process_pdfs_ocr(self):
        """Process PDFs with OCR in separate thread"""
        input_path = self.pdf_input_path.get()
        if not input_path:
            messagebox.showerror("Error", "Please select a PDF file or folder")
            return
        
        if not os.path.exists(input_path):
            messagebox.showerror("Error", "Selected path does not exist")
            return
        
        # Disable controls
        self.toggle_pdf_controls(False)
        self.pdf_log_message("=" * 60)
        self.pdf_log_message("🔄 Starting PDF OCR Processing")
        self.pdf_log_message("=" * 60)
        
        # Start processing in separate thread
        thread = threading.Thread(target=self.run_pdf_processing, args=(input_path,))
        thread.daemon = True
        thread.start()
    
    def run_pdf_processing(self, input_path):
        """Run PDF processing with OCR"""
        try:
            pdf_processor = MultiSubjectPDFProcessor(OCR_API_KEY)
            
            if os.path.isfile(input_path):
                # Process single file
                self.pdf_log_message(f"Processing file: {os.path.basename(input_path)}")
                
                # Extract text with OCR
                text = pdf_processor.extract_text_with_ocr(input_path, self.pdf_log_text)
                
                if text:
                    # Save extracted text as PDF
                    output_filename = f"{os.path.splitext(os.path.basename(input_path))[0]}_extracted.pdf"
                    output_path = os.path.join(OUTPUT_DIR, output_filename)
                    
                    if pdf_processor.create_searchable_pdf(text, output_path):
                        self.pdf_log_message(f"✅ Created searchable PDF: {output_filename}")
                    else:
                        self.pdf_log_message("❌ Failed to create PDF")
            else:
                # Process folder
                self.pdf_log_message(f"Processing folder: {input_path}")
                
                pdf_files = glob.glob(os.path.join(input_path, "*.pdf")) + \
                           glob.glob(os.path.join(input_path, "*.PDF"))
                
                self.pdf_log_message(f"Found {len(pdf_files)} PDF file(s)")
                
                for i, pdf_file in enumerate(pdf_files, 1):
                    self.pdf_log_message(f"\nProcessing {i}/{len(pdf_files)}: {os.path.basename(pdf_file)}")
                    
                    text = pdf_processor.extract_text_with_ocr(pdf_file, self.pdf_log_text)
                    
                    if text:
                        output_filename = f"{os.path.splitext(os.path.basename(pdf_file))[0]}_extracted.pdf"
                        output_path = os.path.join(OUTPUT_DIR, output_filename)
                        
                        if pdf_processor.create_searchable_pdf(text, output_path):
                            self.pdf_log_message(f"✅ Created: {output_filename}")
                        else:
                            self.pdf_log_message("❌ Failed to create PDF")
            
            self.pdf_log_message("\n" + "=" * 60)
            self.pdf_log_message("🎉 PDF Processing Complete!")
            self.pdf_log_message("=" * 60)
            
            # Enable controls
            self.root.after(0, lambda: self.toggle_pdf_controls(True))
            
            # Show success message
            self.root.after(0, lambda: messagebox.showinfo(
                "Success",
                "PDF processing completed successfully!"
            ))
            
        except Exception as e:
            error_msg = str(e)
            self.pdf_log_message(f"\n❌ Processing failed: {error_msg}")
            self.root.after(0, lambda: self.toggle_pdf_controls(True))
            self.root.after(0, lambda: messagebox.showerror("Error", f"Processing failed: {error_msg}"))
    
    def test_email_connection(self):
        """Test email connection"""
        try:
            email_sender = EmailSender(self.sender_email_var.get(), self.app_password_var.get())
            success, message = email_sender.test_connection()
            
            if success:
                messagebox.showinfo("Success", "Email connection successful!")
            else:
                messagebox.showerror("Error", message)
        except Exception as e:
            messagebox.showerror("Error", f"Test failed: {str(e)}")
    
    def save_settings(self):
        """Save system settings"""
        global SENDER_EMAIL, APP_PASSWORD, OCR_API_KEY, OUTPUT_DIR
        
        SENDER_EMAIL = self.sender_email_var.get()
        APP_PASSWORD = self.app_password_var.get()
        OCR_API_KEY = self.ocr_api_key_var.get()
        OUTPUT_DIR = self.output_dir_var.get()
        
        # Create output directory if it doesn't exist
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
        
        messagebox.showinfo("Success", "Settings saved successfully!")
    
    def open_results_folder(self):
        """Open the results folder"""
        if os.path.exists(OUTPUT_DIR):
            os.startfile(OUTPUT_DIR)  # Windows
            # For macOS: os.system(f'open "{OUTPUT_DIR}"')
            # For Linux: os.system(f'xdg-open "{OUTPUT_DIR}"')
        else:
            messagebox.showinfo("Info", "Results folder does not exist yet.")
    
    def clear_log(self):
        """Clear the evaluation log"""
        self.log_text.delete(1.0, tk.END)
    
    def log_message(self, message):
        """Add message to evaluation log"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
    
    def pdf_log_message(self, message):
        """Add message to PDF processing log"""
        self.pdf_log_text.insert(tk.END, message + "\n")
        self.pdf_log_text.see(tk.END)
    
    def toggle_controls(self, enabled):
        """Enable/disable evaluation controls"""
        state = 'normal' if enabled else 'disabled'
        
        # Update status
        if enabled:
            self.status_bar.config(text="Ready")
        else:
            self.status_bar.config(text="Processing...")
    
    def toggle_pdf_controls(self, enabled):
        """Enable/disable PDF processing controls"""
        state = 'normal' if enabled else 'disabled'
        # Note: In a full implementation, you would disable PDF processing controls here

def main():
    """Main function to run the application"""
    root = tk.Tk()
    
    # Configure styles
    style = ttk.Style()
    style.configure('Accent.TButton', font=('Arial', 10, 'bold'))
    
    # Create output directory
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
    
    # Create and run the application
    app = MultiSubjectGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()