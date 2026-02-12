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

# ========================================
# CONFIGURATION
# ========================================
OCR_API_KEY = "K86485627588957"
SENDER_EMAIL = "nitesh.t.mulam2004@gmail.com"
APP_PASSWORD = "gxdd zdyh gfym mlcq"
OUTPUT_DIR = "extracted_pdfs"

class PDFProcessor:
    """Handles PDF processing including OCR extraction"""
    
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
    """Handles email sending functionality"""
    
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
    
    def send_results_email(self, student_data, results_file=None):
        """Send results email to student"""
        try:
            student_name = student_data['Name']
            student_email = student_data['Email']
            roll_no = student_data['Roll No']
            total_marks = student_data['Total Marks']
            percentage = student_data['Percentage']
            grade = student_data['Grade']
            
            # Validate email
            if not student_email or student_email == f"student_{student_data.get('id', 0)}@example.com":
                return False, f"No valid email found for {student_name}"
            
            # Create email message
            message = MIMEMultipart()
            message["From"] = self.sender_email
            message["To"] = student_email
            message["Subject"] = f"📊 Exam Results - {student_name}"
            
            # Create email body
            body = f"""
<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
    <h2 style="color: #2c3e50;">🎓 Exam Results</h2>
    <p>Dear <strong>{student_name}</strong>,</p>
    
    <p>Your exam evaluation is complete. Here are your results:</p>
    
    <div style="background-color: #f8f9fa; padding: 20px; border-radius: 10px; margin: 20px 0;">
        <h3 style="color: #3498db; margin-top: 0;">📊 Result Summary</h3>
        <table style="width: 100%; border-collapse: collapse;">
            <tr>
                <td style="padding: 8px; border-bottom: 1px solid #ddd;"><strong>Roll Number:</strong></td>
                <td style="padding: 8px; border-bottom: 1px solid #ddd;">{roll_no}</td>
            </tr>
            <tr>
                <td style="padding: 8px; border-bottom: 1px solid #ddd;"><strong>Total Marks:</strong></td>
                <td style="padding: 8px; border-bottom: 1px solid #ddd;">{total_marks:.2f}/100</td>
            </tr>
            <tr>
                <td style="padding: 8px; border-bottom: 1px solid #ddd;"><strong>Percentage:</strong></td>
                <td style="padding: 8px; border-bottom: 1px solid #ddd;">{percentage:.2f}%</td>
            </tr>
            <tr>
                <td style="padding: 8px;"><strong>Grade:</strong></td>
                <td style="padding: 8px; font-weight: bold; color: {'#27ae60' if grade in ['A+', 'A', 'B+'] else '#e74c3c'};">{grade}</td>
            </tr>
        </table>
    </div>
    
    <div style="background-color: #e8f4fc; padding: 15px; border-radius: 8px; margin: 20px 0;">
        <h4 style="color: #2980b9; margin-top: 0;">📈 Performance Feedback</h4>
        <p>{self.get_performance_feedback(percentage)}</p>
    </div>
    
    <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee;">
        <p><strong>Note:</strong> For detailed question-wise scores, please check the attached Excel file.</p>
        <p><em>This is an automated message. Please do not reply to this email.</em></p>
    </div>
    
    <div style="margin-top: 20px; color: #7f8c8d; font-size: 12px;">
        <p>Generated on: {datetime.now().strftime("%B %d, %Y at %I:%M %p")}</p>
        <p>© Exam Evaluation System</p>
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
    
    def get_performance_feedback(self, percentage):
        """Generate performance feedback"""
        if percentage >= 90:
            return "🌟 Outstanding performance! Excellent understanding of the subject."
        elif percentage >= 80:
            return "🎯 Excellent work! You have a strong grasp of the material."
        elif percentage >= 70:
            return "👍 Very good performance. Keep up the good work!"
        elif percentage >= 60:
            return "📚 Good effort. Some areas could use improvement."
        elif percentage >= 50:
            return "⚠️ Average performance. Consider reviewing key concepts."
        elif percentage >= 40:
            return "📉 Below average. Please focus on core topics."
        else:
            return "❌ Needs significant improvement. Consider seeking additional help."

class AnswerSheetEvaluator:
    """Main evaluation logic"""
    
    def __init__(self, use_ocr=False):
        self.use_ocr = use_ocr
        self.results_file = None
        self.master_answers = {}
        self.log_messages = []
        self.email_sender = EmailSender(SENDER_EMAIL, APP_PASSWORD)
        self.pdf_processor = PDFProcessor(OCR_API_KEY)
    
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
    
    def evaluate(self, master_pdf_path, student_pdf_paths, log_widget=None, progress_callback=None):
        """Main evaluation workflow"""
        self.log("=" * 60, log_widget)
        self.log("📚 STARTING EVALUATION PROCESS", log_widget)
        self.log("=" * 60, log_widget)
        
        # Create output directory
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
        
        # Step 1: Extract Master Answers
        self.log("\nSTEP 1: EXTRACTING MASTER ANSWERS", log_widget)
        self.log(f"Processing master: {os.path.basename(master_pdf_path)}", log_widget)
        
        master_text = self.extract_text_from_pdf(master_pdf_path, "Master", log_widget)
        if not master_text:
            raise Exception("Failed to extract text from master PDF")
        
        self.master_answers = self.parse_master_answer(master_text)
        if not self.master_answers:
            raise Exception("Could not parse questions from master")
        
        self.log(f"✓ Found {len(self.master_answers)} questions in master", log_widget)
        
        # Step 2: Process Student Answer Sheets
        self.log(f"\nSTEP 2: PROCESSING {len(student_pdf_paths)} STUDENT ANSWER SHEETS", log_widget)
        results = []
        max_score = 10 * len(self.master_answers)
        
        for i, student_pdf in enumerate(student_pdf_paths):
            filename = os.path.basename(student_pdf)
            self.log(f"\n🔍 Processing {filename} ({i+1}/{len(student_pdf_paths)})", log_widget)
            
            if progress_callback:
                progress_callback((i + 1) / len(student_pdf_paths) * 50 + 25)
            
            student_text = self.extract_text_from_pdf(student_pdf, filename, log_widget)
            if not student_text:
                self.log(f"  ✗ Failed to extract text", log_widget)
                continue
            
            student_data = self.parse_student_answers(student_text, i+1)
            
            # Evaluate answers
            total_score = 0
            question_scores = {}
            
            for q_key, m_ans in self.master_answers.items():
                s_ans = student_data['answers'].get(q_key, '')
                score = self.evaluate_answer(m_ans, s_ans)
                question_scores[q_key] = score
                total_score += score
            
            percentage = round((total_score / max_score) * 100, 2) if max_score > 0 else 0
            grade = self.calculate_grade(percentage)
            
            result = {
                'ID': i+1,
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
            self.log(f"  📧 Email: {student_data['email']}", log_widget)
            self.log(f"  📊 Score: {total_score:.1f}/{max_score} ({percentage}%)", log_widget)
            self.log(f"  🏆 Grade: {grade}", log_widget)
        
        # Step 3: Save Results
        self.log("\nSTEP 3: SAVING RESULTS", log_widget)
        
        if results:
            self.results_file = self.save_to_excel(results)
            self.log(f"✓ Results saved to: {self.results_file}", log_widget)
            
            # Display summary
            self.display_summary(results, log_widget)
            
            if progress_callback:
                progress_callback(100)
        else:
            raise Exception("No student results to evaluate")
        
        return results
    
    def send_emails(self, results, log_widget=None, progress_callback=None):
        """Send emails to all students"""
        self.log("\n" + "=" * 60, log_widget)
        self.log("📧 SENDING EMAILS TO STUDENTS", log_widget)
        self.log("=" * 60, log_widget)
        
        # Test email connection
        self.log("Testing email connection...", log_widget)
        success, message = self.email_sender.test_connection()
        if not success:
            self.log(f"✗ {message}", log_widget)
            return 0, len(results)
        
        self.log(f"✓ {message}", log_widget)
        
        success_count = 0
        fail_count = 0
        
        for i, result in enumerate(results):
            self.log(f"\n📨 Sending to {result['Name']} ({i+1}/{len(results)})", log_widget)
            self.log(f"  Email: {result['Email']}", log_widget)
            
            if progress_callback:
                progress_callback((i + 1) / len(results) * 100)
            
            success, msg = self.email_sender.send_results_email(result, self.results_file)
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
    
    def display_summary(self, results, log_widget=None):
        """Display summary statistics"""
        self.log("\n" + "=" * 60, log_widget)
        self.log("📊 EVALUATION SUMMARY", log_widget)
        self.log("=" * 60, log_widget)
        
        if not results:
            self.log("No results to display", log_widget)
            return
        
        percentages = [r['Percentage'] for r in results]
        
        self.log(f"Total Students Evaluated: {len(results)}", log_widget)
        self.log(f"Average Score: {sum(percentages)/len(percentages):.1f}%", log_widget)
        self.log(f"Highest Score: {max(percentages):.1f}%", log_widget)
        self.log(f"Lowest Score: {min(percentages):.1f}%", log_widget)
        
        # Grade distribution
        grade_dist = {}
        for result in results:
            grade = result['Grade']
            grade_dist[grade] = grade_dist.get(grade, 0) + 1
        
        self.log("\nGrade Distribution:", log_widget)
        for grade, count in sorted(grade_dist.items(), key=lambda x: x[0]):
            self.log(f"  {grade}: {count} student(s)", log_widget)
    
    def save_to_excel(self, results):
        """Save results to Excel"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(OUTPUT_DIR, f"evaluation_results_{timestamp}.xlsx")
        
        df = pd.DataFrame(results)
        
        # Reorder columns
        base_cols = ['ID', 'Name', 'Roll No', 'Email', 'Total Marks', 'Percentage', 'Grade']
        question_cols = [col for col in df.columns if col not in base_cols]
        ordered_cols = base_cols + sorted(question_cols)
        
        for col in ordered_cols:
            if col not in df.columns:
                df[col] = ''
        
        df = df[ordered_cols]
        
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
        
        return filename

class AIAnswerSheetEvaluatorGUI:
    """GUI Application for AI-based Answer Sheet Evaluation"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("AI Answer Sheet Evaluation System")
        self.root.geometry("1000x700")
        
        # Variables
        self.master_pdf_path = tk.StringVar()
        self.student_pdfs = []
        self.use_ocr = tk.BooleanVar(value=True)
        self.send_emails = tk.BooleanVar(value=False)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_evaluation_tab()
        self.create_pdf_processing_tab()
        self.create_settings_tab()
        
        # Status bar
        self.status_bar = tk.Label(root, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def create_evaluation_tab(self):
        """Create the evaluation tab"""
        self.eval_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.eval_frame, text="Evaluation")
        
        # Title
        title_label = tk.Label(self.eval_frame, text="🎓 AI Answer Sheet Evaluation System", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Main content frame
        content_frame = ttk.Frame(self.eval_frame)
        content_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Left panel for controls
        left_panel = ttk.Frame(content_frame)
        left_panel.pack(side='left', fill='y', padx=(0, 10))
        
        # Master PDF Selection
        master_frame = ttk.LabelFrame(left_panel, text="Master Answer Sheet", padding=10)
        master_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(master_frame, text="Select Master PDF:").pack(anchor='w')
        
        master_path_frame = ttk.Frame(master_frame)
        master_path_frame.pack(fill='x', pady=5)
        
        ttk.Entry(master_path_frame, textvariable=self.master_pdf_path, width=40).pack(side='left', fill='x', expand=True)
        ttk.Button(master_path_frame, text="Browse", command=self.browse_master_pdf).pack(side='left', padx=5)
        
        # Student PDFs Selection
        student_frame = ttk.LabelFrame(left_panel, text="Student Answer Sheets", padding=10)
        student_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(student_frame, text="Select Student PDFs:").pack(anchor='w')
        
        student_buttons_frame = ttk.Frame(student_frame)
        student_buttons_frame.pack(fill='x', pady=5)
        
        ttk.Button(student_buttons_frame, text="Browse Files", command=self.browse_student_pdfs).pack(side='left', padx=2)
        ttk.Button(student_buttons_frame, text="Browse Folder", command=self.browse_student_folder).pack(side='left', padx=2)
        ttk.Button(student_buttons_frame, text="Clear All", command=self.clear_student_pdfs).pack(side='left', padx=2)
        
        # Student listbox
        self.student_listbox = tk.Listbox(student_frame, height=8)
        self.student_listbox.pack(fill='both', expand=True, pady=5)
        
        # Options frame
        options_frame = ttk.LabelFrame(left_panel, text="Evaluation Options", padding=10)
        options_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Checkbutton(options_frame, text="Use OCR (for handwritten sheets)", 
                       variable=self.use_ocr).pack(anchor='w', pady=2)
        ttk.Checkbutton(options_frame, text="Send Emails to Students", 
                       variable=self.send_emails).pack(anchor='w', pady=2)
        
        # Control buttons
        control_frame = ttk.Frame(left_panel)
        control_frame.pack(fill='x')
        
        ttk.Button(control_frame, text="📊 Start Evaluation", 
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
    
    def browse_master_pdf(self):
        """Browse for master PDF file"""
        file_path = filedialog.askopenfilename(
            title="Select Master Answer Sheet",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.master_pdf_path.set(file_path)
            self.log_message(f"Selected master PDF: {os.path.basename(file_path)}")
    
    def browse_student_pdfs(self):
        """Browse for student PDF files"""
        file_paths = filedialog.askopenfilenames(
            title="Select Student Answer Sheets",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        for file_path in file_paths:
            if file_path not in self.student_pdfs:
                self.student_pdfs.append(file_path)
                self.student_listbox.insert(tk.END, os.path.basename(file_path))
        self.log_message(f"Added {len(file_paths)} student PDF(s)")
    
    def browse_student_folder(self):
        """Browse for folder containing student PDFs"""
        folder_path = filedialog.askdirectory(title="Select Folder with Student Answer Sheets")
        if folder_path:
            pdf_files = glob.glob(os.path.join(folder_path, "*.pdf")) + \
                       glob.glob(os.path.join(folder_path, "*.PDF"))
            for file_path in pdf_files:
                if file_path not in self.student_pdfs:
                    self.student_pdfs.append(file_path)
                    self.student_listbox.insert(tk.END, os.path.basename(file_path))
            self.log_message(f"Added {len(pdf_files)} student PDF(s) from folder")
    
    def clear_student_pdfs(self):
        """Clear all selected student PDFs"""
        self.student_pdfs.clear()
        self.student_listbox.delete(0, tk.END)
        self.log_message("Cleared all student PDFs")
    
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
    
    def start_evaluation(self):
        """Start the evaluation process in a separate thread"""
        if not self.master_pdf_path.get():
            messagebox.showerror("Error", "Please select a master answer sheet PDF")
            return
        
        if not self.student_pdfs:
            messagebox.showerror("Error", "Please select at least one student answer sheet PDF")
            return
        
        if not os.path.exists(self.master_pdf_path.get()):
            messagebox.showerror("Error", "Master PDF file does not exist")
            return
        
        # Validate student files exist
        valid_files = []
        for file_path in self.student_pdfs:
            if os.path.exists(file_path):
                valid_files.append(file_path)
            else:
                self.log_message(f"Warning: File not found: {file_path}")
        
        if not valid_files:
            messagebox.showerror("Error", "No valid student PDF files found")
            return
        
        # Disable controls during evaluation
        self.toggle_controls(False)
        self.progress_bar['value'] = 0
        self.log_message("=" * 60)
        self.log_message("🚀 Starting Evaluation Process")
        self.log_message(f"Master PDF: {os.path.basename(self.master_pdf_path.get())}")
        self.log_message(f"Student PDFs: {len(valid_files)} files")
        self.log_message(f"OCR Enabled: {self.use_ocr.get()}")
        self.log_message("=" * 60)
        
        # Start evaluation in separate thread
        thread = threading.Thread(target=self.run_evaluation, args=(valid_files,))
        thread.daemon = True
        thread.start()
    
    def run_evaluation(self, student_files):
        """Run the evaluation process"""
        try:
            # Create evaluator instance
            evaluator = AnswerSheetEvaluator(use_ocr=self.use_ocr.get())
            
            # Run evaluation with progress updates
            def update_progress(value):
                self.root.after(0, lambda: self.progress_bar.config(value=value))
            
            results = evaluator.evaluate(
                self.master_pdf_path.get(),
                student_files,
                self.log_text,
                update_progress
            )
            
            # Send emails if requested
            if self.send_emails.get():
                success_count, fail_count = evaluator.send_emails(
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
                f"Results saved to: {evaluator.results_file}"
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
        # This would require loading existing results
        messagebox.showinfo("Info", "This feature requires an existing results file.")
    
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
            pdf_processor = PDFProcessor(OCR_API_KEY)
            
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
    app = AIAnswerSheetEvaluatorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()