import requests
import json
import base64
import os
import re
import pandas as pd
import nltk
from PyPDF2 import PdfReader
from datetime import datetime
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import argparse
import sys

# ========================================
# CONFIGURATION
# ========================================
OCR_API_KEY = "K86485627588957"
SENDER_EMAIL = "nitesh.t.mulam2004@gmail.com"
APP_PASSWORD = "gxdd zdyh gfym mlcq"
OUTPUT_DIR = "extracted_pdfs"

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
            
            # Attach HTML body
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
                    print(f"  ⚠️ Could not attach file for {student_name}: {e}")
            
            # Send email
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(self.sender_email, self.app_password)
                server.sendmail(self.sender_email, student_email, message.as_string())
            
            # Log successful email
            self.sent_emails_log.append({
                'student': student_name,
                'email': student_email,
                'time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'status': 'Sent'
            })
            
            return True, f"Email sent to {student_name} ({student_email})"
            
        except Exception as e:
            error_msg = str(e)
            # Log failed email
            self.sent_emails_log.append({
                'student': student_name,
                'email': student_email,
                'time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'status': 'Failed',
                'error': error_msg
            })
            return False, f"Failed to send email to {student_name}: {error_msg}"
    
    def get_performance_feedback(self, percentage):
        """Generate performance feedback based on percentage"""
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
    """Main evaluation logic for command line"""
    
    def __init__(self, use_ocr=False):
        self.use_ocr = use_ocr
        self.results_file = None
        self.master_answers = {}
        self.log_messages = []
        self.email_sender = EmailSender(SENDER_EMAIL, APP_PASSWORD)
    
    def log(self, message):
        """Log message"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        print(log_entry)
        self.log_messages.append(log_entry)
    
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
    
    def extract_pdf_ocr(self, pdf_path, source_name):
        """Extract text using OCR"""
        try:
            self.log(f"  🔄 Using OCR extraction...")
            
            with open(pdf_path, "rb") as file:
                pdf_bytes = file.read()
            
            pdf_base64 = base64.b64encode(pdf_bytes).decode('utf-8')
            
            url = "https://api.ocr.space/parse/image"
            data = {
                "apikey": OCR_API_KEY,
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
            
            if result["IsErroredOnProcessing"]:
                return None
            
            extracted_text = ""
            parsed_results = result.get("ParsedResults", [])
            
            for page in parsed_results:
                page_text = page.get("ParsedText", "")
                extracted_text += page_text + "\n"
            
            return extracted_text
        except Exception as e:
            self.log(f"  ✗ OCR extraction failed: {e}")
            return None
    
    def clean_extracted_text(self, text):
        """Clean extracted text"""
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'(Q\d+)', r'\n\1:', text)
        return text.strip()
    
    def extract_master_text(self, pdf_path):
        """Extract text from master PDF"""
        try:
            text = self.extract_pdf_text(pdf_path)
            if text and len(text.strip()) > 100:
                return self.clean_extracted_text(text)
            
            if self.use_ocr and OCR_API_KEY:
                ocr_text = self.extract_pdf_ocr(pdf_path, "Master")
                if ocr_text:
                    return self.clean_extracted_text(ocr_text)
            
            return text
        except Exception as e:
            self.log(f"✗ Error extracting master text: {e}")
            return None
    
    def extract_student_text(self, pdf_path, student_name):
        """Extract text from student PDF"""
        try:
            text = self.extract_pdf_text(pdf_path)
            if text and len(text.strip()) > 50:
                return self.clean_extracted_text(text)
            
            if self.use_ocr and OCR_API_KEY:
                ocr_text = self.extract_pdf_ocr(pdf_path, student_name)
                if ocr_text:
                    return self.clean_extracted_text(ocr_text)
            
            return text
        except Exception as e:
            self.log(f"  ✗ Error extracting student text: {e}")
            return None
    
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
        
        # Extract name (multiple patterns)
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
        
        # Extract email (multiple patterns)
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
        """Parse student data including EMAIL"""
        # Extract student info
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
        """Evaluate a single answer"""
        if not student_answer or student_answer.strip() == "":
            return 0.0
        
        try:
            # Preprocess text
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
    
    def evaluate(self, master_pdf_path, student_pdf_paths):
        """Main evaluation workflow"""
        self.log("=" * 60)
        self.log("📚 STARTING EVALUATION PROCESS")
        self.log("=" * 60)
        
        # Create output directory
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
        
        # Step 1: Extract Master Answers
        self.log("\nSTEP 1: EXTRACTING MASTER ANSWERS")
        self.log(f"Processing master: {os.path.basename(master_pdf_path)}")
        
        master_text = self.extract_master_text(master_pdf_path)
        if not master_text:
            raise Exception("Failed to extract text from master PDF")
        
        self.master_answers = self.parse_master_answer(master_text)
        if not self.master_answers:
            raise Exception("Could not parse questions from master")
        
        self.log(f"✓ Found {len(self.master_answers)} questions in master")
        
        # Step 2: Process Student Answer Sheets
        self.log(f"\nSTEP 2: PROCESSING {len(student_pdf_paths)} STUDENT ANSWER SHEETS")
        results = []
        max_score = 10 * len(self.master_answers)
        
        for i, student_pdf in enumerate(student_pdf_paths):
            filename = os.path.basename(student_pdf)
            self.log(f"\n🔍 Processing {filename} ({i+1}/{len(student_pdf_paths)})")
            
            student_text = self.extract_student_text(student_pdf, filename)
            if not student_text:
                self.log(f"  ✗ Failed to extract text")
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
            
            self.log(f"  👤 Name: {student_data['name']}")
            self.log(f"  📝 Roll No: {student_data['roll_no']}")
            self.log(f"  📧 Email: {student_data['email']}")
            self.log(f"  📊 Score: {total_score:.1f}/{max_score} ({percentage}%)")
            self.log(f"  🏆 Grade: {grade}")
        
        # Step 3: Save Results
        self.log("\nSTEP 3: SAVING RESULTS")
        
        if results:
            self.results_file = self.save_to_excel(results)
            self.log(f"✓ Results saved to: {self.results_file}")
            
            # Display summary
            self.display_summary(results)
        else:
            raise Exception("No student results to evaluate")
        
        return results
    
    def send_emails(self, results):
        """Send emails to all students with their results"""
        self.log("\n" + "=" * 60)
        self.log("📧 SENDING EMAILS TO STUDENTS")
        self.log("=" * 60)
        
        # Test email connection first
        self.log("Testing email connection...")
        success, message = self.email_sender.test_connection()
        if not success:
            self.log(f"✗ {message}")
            return 0, len(results)
        
        self.log(f"✓ {message}")
        
        success_count = 0
        fail_count = 0
        
        for i, result in enumerate(results):
            self.log(f"\n📨 Sending to {result['Name']} ({i+1}/{len(results)})")
            self.log(f"  Email: {result['Email']}")
            
            success, message = self.email_sender.send_results_email(result, self.results_file)
            if success:
                self.log(f"  ✅ {message}")
                success_count += 1
            else:
                self.log(f"  ❌ {message}")
                fail_count += 1
        
        # Save email log
        self.save_email_log()
        
        return success_count, fail_count
    
    def save_email_log(self):
        """Save email sending log"""
        if self.email_sender.sent_emails_log:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_file = os.path.join(OUTPUT_DIR, f"email_log_{timestamp}.csv")
            
            df = pd.DataFrame(self.email_sender.sent_emails_log)
            df.to_csv(log_file, index=False)
            self.log(f"\n📝 Email log saved to: {log_file}")
    
    def display_summary(self, results):
        """Display summary statistics"""
        self.log("\n" + "=" * 60)
        self.log("📊 EVALUATION SUMMARY")
        self.log("=" * 60)
        
        if not results:
            self.log("No results to display")
            return
        
        percentages = [r['Percentage'] for r in results]
        
        self.log(f"Total Students Evaluated: {len(results)}")
        self.log(f"Average Score: {sum(percentages)/len(percentages):.1f}%")
        self.log(f"Highest Score: {max(percentages):.1f}%")
        self.log(f"Lowest Score: {min(percentages):.1f}%")
        
        # Display grade distribution
        grade_dist = {}
        for result in results:
            grade = result['Grade']
            grade_dist[grade] = grade_dist.get(grade, 0) + 1
        
        self.log("\nGrade Distribution:")
        for grade, count in sorted(grade_dist.items(), key=lambda x: x[0]):
            self.log(f"  {grade}: {count} student(s)")
    
    def save_to_excel(self, results):
        """Save results to Excel"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(OUTPUT_DIR, f"evaluation_results_{timestamp}.xlsx")
        
        df = pd.DataFrame(results)
        
        # Reorder columns
        base_cols = ['ID', 'Name', 'Roll No', 'Email', 'Total Marks', 'Percentage', 'Grade']
        question_cols = [col for col in df.columns if col not in base_cols]
        ordered_cols = base_cols + sorted(question_cols)
        
        # Ensure all columns exist
        for col in ordered_cols:
            if col not in df.columns:
                df[col] = ''
        
        df = df[ordered_cols]
        
        # Save to Excel
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Results', index=False)
            
            # Auto-adjust column widths
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

def main():
    """Main function to run the command-line application"""
    
    parser = argparse.ArgumentParser(
        description='Answer Sheet Evaluator - Automated OCR Extraction & NLP Evaluation System'
    )
    
    parser.add_argument('--master', required=True, help='Path to master answer sheet PDF')
    parser.add_argument('--students', nargs='+', help='Paths to student answer sheet PDFs')
    parser.add_argument('--folder', help='Folder containing student PDFs')
    parser.add_argument('--ocr', action='store_true', help='Use OCR for handwritten sheets')
    parser.add_argument('--send-emails', action='store_true', help='Send emails to students')
    parser.add_argument('--output', default='extracted_pdfs', help='Output directory for results')
    parser.add_argument('--test-email', action='store_true', help='Test email connection only')
    
    args = parser.parse_args()
    
    # Update OUTPUT_DIR if specified
    global OUTPUT_DIR
    if args.output:
        OUTPUT_DIR = args.output
    
    # Test email only
    if args.test_email:
        email_sender = EmailSender(SENDER_EMAIL, APP_PASSWORD)
        success, message = email_sender.test_connection()
        if success:
            print(f"✅ {message}")
            # Send test email
            test_data = {
                'Name': 'Test Student',
                'Roll No': 'TEST001',
                'Email': SENDER_EMAIL,  # Send to yourself
                'Total Marks': 85.5,
                'Percentage': 85.5,
                'Grade': 'A'
            }
            success2, message2 = email_sender.send_results_email(test_data)
            print(f"📧 Test email: {message2}")
        else:
            print(f"❌ {message}")
        sys.exit(0)
    
    # Check for required files
    if not os.path.exists(args.master):
        print(f"Error: Master PDF not found: {args.master}")
        sys.exit(1)
    
    # Get student files
    student_files = []
    
    if args.students:
        student_files = args.students
    elif args.folder:
        if not os.path.exists(args.folder):
            print(f"Error: Folder not found: {args.folder}")
            sys.exit(1)
        student_files = [os.path.join(args.folder, f) for f in os.listdir(args.folder) 
                        if f.lower().endswith('.pdf')]
    
    if not student_files:
        print("Error: No student PDF files provided. Use --students or --folder")
        sys.exit(1)
    
    # Validate student files exist
    valid_student_files = []
    for file in student_files:
        if os.path.exists(file):
            valid_student_files.append(file)
        else:
            print(f"Warning: Student PDF not found: {file}")
    
    if not valid_student_files:
        print("Error: No valid student PDF files found")
        sys.exit(1)
    
    print("\n" + "=" * 60)
    print("ANSWER SHEET EVALUATOR - COMMAND LINE VERSION")
    print("=" * 60)
    print(f"Master PDF: {os.path.basename(args.master)}")
    print(f"Student PDFs: {len(valid_student_files)} files")
    print(f"OCR Enabled: {args.ocr}")
    print(f"Email Sending: {args.send_emails}")
    print("=" * 60 + "\n")
    
    try:
        # Create evaluator instance
        evaluator = AnswerSheetEvaluator(use_ocr=args.ocr)
        
        # Run evaluation
        results = evaluator.evaluate(args.master, valid_student_files)
        
        # Send emails if requested
        if args.send_emails:
            success_count, fail_count = evaluator.send_emails(results)
            print(f"\n📧 Email Summary:")
            print(f"  ✅ Successfully sent: {success_count}")
            print(f"  ❌ Failed to send: {fail_count}")
            
            if fail_count > 0:
                print("\n⚠️ Some emails failed to send. Check the email log for details.")
        
        print("\n" + "=" * 60)
        print("✅ EVALUATION COMPLETED SUCCESSFULLY")
        print("=" * 60)
        print(f"Results saved to: {evaluator.results_file}")
        print("=" * 60)
        
        # Ask to view results
        choice = input("\nWould you like to view the results table? (y/n): ").lower()
        if choice == 'y':
            print("\n" + "=" * 80)
            print(f"{'Name':<20} {'Roll No':<10} {'Email':<25} {'Score':<8} {'Grade':<5}")
            print("=" * 80)
            for result in results:
                print(f"{result['Name'][:20]:<20} {result['Roll No']:<10} "
                      f"{result['Email'][:25]:<25} {result['Percentage']:>6.1f}% {result['Grade']:>5}")
            print("=" * 80)
        
    except Exception as e:
        print(f"\n❌ Evaluation failed: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()