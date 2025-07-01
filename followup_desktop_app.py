# desktop_email_followup_app.py - Desktop Email Follow-up Application

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import queue
import logging
from datetime import datetime, timedelta
import json
import re
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import ssl
import pytz
from exchangelib import DELEGATE, Account, Credentials, Configuration, Message, EWSTimeZone, Folder
from bs4 import BeautifulSoup

class EmailFollowupApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Follow-up Desktop Application - Mohammad Reza Kazemi")
        self.root.geometry("1200x800")
        
        # Configure logging
        self.setup_logging()
        
        # Application state
        self.conversations = {}
        self.followup_candidates = []
        self.user_email = ""
        self.user_password = ""
        self.is_processing = False
        
        # Queue for thread communication
        self.message_queue = queue.Queue()
        
        # LLM Configuration
        self.LLM_API_ENDPOINT = "https://api.openai.com/v1/chat/completions"
        self.LLM_MODEL_NAME = "gpt-3.5-turbo"
        self.LLM_API_KEY = "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"
        self.LLM_SIMULATION = False
        
        # SMTP Configuration
        self.SMTP_SERVER = "mail.mtnirancell.ir"
        self.SMTP_PORT = 587
        
        self.setup_ui()
        self.check_message_queue()

    def send_email_ews(self, subject, body, to_addresses):
        try:
            # Create credentials and account (reuse existing connection if possible)
            credentials = Credentials(self.user_email, self.user_password)
            account = Account(primary_smtp_address=self.user_email, credentials=credentials,
                            autodiscover=True, access_type=DELEGATE)
            
            # Create message
            message = Message(
                account=account,
                subject=subject,
                body=body,
                to_recipients=to_addresses
            )
            
            # Send and save to Sent Items
            message.send(save_copy=True, copy_to_folder=account.sent)
            
            return True
            
        except Exception as e:
            self.message_queue.put({'type': 'log', 'text': f'EWS email sending error: {str(e)}', 'level': 'ERROR'})
            return False
        
    def setup_logging(self):
        """Setup logging configuration"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('email_followup_logs.txt'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
    def setup_ui(self):
        """Setup the main user interface"""
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Email Follow-up Application", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Credentials section
        cred_frame = ttk.LabelFrame(main_frame, text="Email Credentials", padding="10")
        cred_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        cred_frame.columnconfigure(1, weight=1)
        
        ttk.Label(cred_frame, text="Email:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.email_var = tk.StringVar()
        self.email_entry = ttk.Entry(cred_frame, textvariable=self.email_var, width=50)
        self.email_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Label(cred_frame, text="Password:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(cred_frame, textvariable=self.password_var, show="*", width=50)
        self.password_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # Number of days to check
        ttk.Label(cred_frame, text="Days to Check:").grid(row=2, column=0, sticky=tk.W, padx=(0, 10))
        self.days_var = tk.IntVar(value=1)
        self.days_spinbox = ttk.Spinbox(cred_frame, from_=1, to=30, textvariable=self.days_var, width=10)
        self.days_spinbox.grid(row=2, column=1, sticky=tk.W, padx=(0, 10))
        
        # Control buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=10)
        
        self.fetch_button = ttk.Button(button_frame, text="Fetch & Analyze Emails", 
                                      command=self.start_analysis)
        self.fetch_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.send_button = ttk.Button(button_frame, text="Send Follow-ups", 
                                     command=self.send_followups, state='disabled')
        self.send_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.export_logs_button = ttk.Button(button_frame, text="Export Logs", 
                                           command=self.export_logs)
        self.export_logs_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Progress bar
        self.progress_var = tk.StringVar()
        self.progress_var.set("Ready")
        self.progress_label = ttk.Label(main_frame, textvariable=self.progress_var)
        self.progress_label.grid(row=4, column=0, columnspan=3, pady=5)
        
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # Analysis results tab
        self.results_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.results_frame, text="Analysis Results")
        self.setup_results_tab()
        
        # Logs tab
        self.logs_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.logs_frame, text="Activity Logs")
        self.setup_logs_tab()
        
        # Follow-up emails tab
        self.emails_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.emails_frame, text="Follow-up Emails")
        self.setup_emails_tab()
        
    def setup_results_tab(self):
        """Setup the analysis results tab"""
        self.results_frame.columnconfigure(0, weight=1)
        self.results_frame.rowconfigure(0, weight=1)
        
        # Treeview for conversations
        columns = ('Subject', 'Main Topic', 'Emails Count', 'Participants', 'Follow-up Needed', 'Reason')
        self.results_tree = ttk.Treeview(self.results_frame, columns=columns, show='tree headings')
        
        # Configure columns
        self.results_tree.heading('#0', text='ID')
        self.results_tree.column('#0', width=50)
        
        for col in columns:
            self.results_tree.heading(col, text=col)
            if col == 'Subject':
                self.results_tree.column(col, width=200)
            elif col == 'Main Topic':
                self.results_tree.column(col, width=250)
            elif col == 'Reason':
                self.results_tree.column(col, width=500)
            else:
                self.results_tree.column(col, width=150)
        
        # Scrollbars
        results_scrollbar_v = ttk.Scrollbar(self.results_frame, orient="vertical", command=self.results_tree.yview)
        results_scrollbar_h = ttk.Scrollbar(self.results_frame, orient="horizontal", command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=results_scrollbar_v.set, xscrollcommand=results_scrollbar_h.set)
        
        self.results_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        results_scrollbar_v.grid(row=0, column=1, sticky=(tk.N, tk.S))
        results_scrollbar_h.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
    def setup_logs_tab(self):
        """Setup the logs tab"""
        self.logs_frame.columnconfigure(0, weight=1)
        self.logs_frame.rowconfigure(0, weight=1)
        
        self.logs_text = scrolledtext.ScrolledText(self.logs_frame, width=80, height=20)
        self.logs_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
    def setup_emails_tab(self):
        """Setup the follow-up emails tab"""
        self.emails_frame.columnconfigure(0, weight=1)
        self.emails_frame.rowconfigure(0, weight=1)
        
        # Create paned window for email list and preview
        paned_window = ttk.PanedWindow(self.emails_frame, orient=tk.HORIZONTAL)
        paned_window.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # Left frame - email list
        left_frame = ttk.Frame(paned_window)
        paned_window.add(left_frame, weight=1)
        
        # Email list
        self.emails_listbox = tk.Listbox(left_frame, width=40)
        self.emails_listbox.pack(fill=tk.BOTH, expand=True)
        self.emails_listbox.bind('<<ListboxSelect>>', self.on_email_select)
        
        # Right frame - email preview
        right_frame = ttk.Frame(paned_window)
        paned_window.add(right_frame, weight=2)
        
        # Email preview
        self.email_preview = scrolledtext.ScrolledText(right_frame, width=60, height=20)
        self.email_preview.pack(fill=tk.BOTH, expand=True)
        
        # Chat box for revising email
        chat_frame = ttk.Frame(right_frame)
        chat_frame.pack(fill=tk.X, pady=(5, 0))
        self.chat_var = tk.StringVar()
        self.chat_entry = ttk.Entry(chat_frame, textvariable=self.chat_var, width=50)
        self.chat_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.chat_button = ttk.Button(chat_frame, text="Revise Email", command=self.revise_email)
        self.chat_button.pack(side=tk.LEFT)
        
    def log_message(self, message, level='INFO'):
        """Add message to logs"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {level}: {message}\n"
        
        self.logs_text.insert(tk.END, log_entry)
        self.logs_text.see(tk.END)
        
        if level == 'INFO':
            self.logger.info(message)
        elif level == 'ERROR':
            self.logger.error(message)
        elif level == 'WARNING':
            self.logger.warning(message)
            
    def check_message_queue(self):
        """Check for messages from worker threads"""
        try:
            while True:
                message = self.message_queue.get_nowait()
                if message['type'] == 'log':
                    self.log_message(message['text'], message['level'])
                elif message['type'] == 'progress':
                    self.progress_var.set(message['text'])
                elif message['type'] == 'complete':
                    self.analysis_complete()
                elif message['type'] == 'error':
                    self.handle_error(message['text'])
        except queue.Empty:
            pass
        
        self.root.after(100, self.check_message_queue)
        
    def start_analysis(self):
        """Start the email analysis in a separate thread"""
        if self.is_processing:
            return
            
        self.user_email = self.email_var.get().strip()
        self.user_password = self.password_var.get().strip()
        self.days_to_check = self.days_var.get()
        
        if not self.user_email or not self.user_password:
            messagebox.showerror("Error", "Please enter both email and password")
            return
            
        self.is_processing = True
        self.fetch_button.config(state='disabled')
        self.send_button.config(state='disabled')
        self.progress_bar.start()
        
        # Clear previous results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        self.emails_listbox.delete(0, tk.END)
        self.email_preview.delete(1.0, tk.END)
        
        # Start analysis thread
        thread = threading.Thread(target=self.analyze_emails_thread)
        thread.daemon = True
        thread.start()
        
    def analyze_emails_thread(self):
        """Thread function for email analysis"""
        try:
            self.message_queue.put({'type': 'log', 'text': 'Starting email analysis...', 'level': 'INFO'})
            self.message_queue.put({'type': 'progress', 'text': 'Connecting to email server...'})
            
            # Fetch emails
            conversations = self.fetch_emails_from_previous_day_ews(self.days_to_check)
            if not conversations:
                self.message_queue.put({'type': 'error', 'text': 'No emails found or connection failed'})
                return
                
            self.conversations = conversations
            self.message_queue.put({'type': 'log', 'text': f'Found {len(conversations)} conversations', 'level': 'INFO'})
            
            # Analyze conversations
            self.message_queue.put({'type': 'progress', 'text': 'Analyzing conversations...'})
            self.followup_candidates = []
            
            for i, (subject, emails) in enumerate(conversations.items()):
                self.message_queue.put({'type': 'progress', 'text': f'Analyzing conversation {i+1}/{len(conversations)}: {subject[:50]}...'})
                
                # Reconstruct conversation text
                conversation_text = ""
                participants = set()
                
                for email in emails:
                    conversation_text += f"{email['sender_name_display']}: {email['content']}\n"
                    participants.add(email['sender'])
                
                # Remove user's own email from participants for follow-up
                participants.discard(self.user_email)
                
                # Generate main topic/summary
                summary_prompt = f"Summarize the main topic of this conversation in one or two sentences.\n\nConversation:\n{conversation_text}"
                main_topic = self.call_llm(summary_prompt)
                
                # Determine if follow-up is needed
                needs_followup, explanation = self.determine_followup_need(conversation_text)
                
                # Log the decision
                self.message_queue.put({'type': 'log', 'text': f'Conversation "{subject}": Follow-up {"NEEDED" if needs_followup else "NOT NEEDED"} - {explanation}', 'level': 'INFO'})
                
                conversation_data = {
                    'subject': subject,
                    'emails': emails,
                    'conversation_text': conversation_text,
                    'participants': list(participants),
                    'explanation': explanation,
                    'needs_followup': needs_followup,
                    'main_topic': main_topic
                }
                
                if needs_followup and participants:  # Only add if there are participants to send to
                    self.followup_candidates.append(conversation_data)
                    
            self.message_queue.put({'type': 'complete', 'text': 'Analysis complete'})
            
        except Exception as e:
            self.message_queue.put({'type': 'error', 'text': f'Analysis failed: {str(e)}'})
            
    def analysis_complete(self):
        """Called when analysis is complete"""
        self.is_processing = False
        self.progress_bar.stop()
        self.fetch_button.config(state='normal')
        self.progress_var.set(f"Analysis complete - {len(self.followup_candidates)} follow-ups needed")
        
        # Populate results tree
        for i, conv in enumerate(self.conversations.values()):
            # Find if this conversation needs follow-up
            needs_followup = False
            explanation = "No follow-up needed"
            main_topic = ""
            for candidate in self.followup_candidates:
                if candidate['subject'] == list(self.conversations.keys())[i]:
                    needs_followup = True
                    explanation = candidate['explanation']
                    main_topic = candidate.get('main_topic', '')
                    break
            participants = set()
            for email in conv:
                participants.add(email['sender'])
            participants.discard(self.user_email)
            self.results_tree.insert('', tk.END, 
                                   text=str(i+1),
                                   values=(
                                       list(self.conversations.keys())[i],
                                       main_topic[:120] + '...' if len(main_topic) > 120 else main_topic,
                                       len(conv),
                                       ', '.join(list(participants)[:3]),  # Show first 3 participants
                                       'YES' if needs_followup else 'NO',
                                       explanation[:100] + '...' if len(explanation) > 100 else explanation
                                   ))
        
        # Generate and populate follow-up emails
        if self.followup_candidates:
            self.send_button.config(state='normal')
            self.generate_followup_emails()
        
        self.log_message(f"Analysis completed. {len(self.followup_candidates)} conversations need follow-up.")
        
    def generate_followup_emails(self):
        """Generate follow-up emails for candidates"""
        self.emails_listbox.delete(0, tk.END)
        
        for i, candidate in enumerate(self.followup_candidates):
            try:
                # Generate follow-up email content
                email_content = self.generate_followup_email_content(candidate)
                candidate['generated_email'] = email_content
                
                # Add to listbox
                self.emails_listbox.insert(tk.END, f"{i+1}. {candidate['subject'][:50]}...")
                
            except Exception as e:
                self.log_message(f"Failed to generate email for '{candidate['subject']}': {str(e)}", 'ERROR')
                
    def generate_followup_email_content(self, candidate):
        """Generate follow-up email content using LLM"""
        # Extract action items
        action_prompt = f"""You are an AI assistant designed to identify action items and the person responsible for each from a conversation. For each action, also note any mentioned deadlines. Output your findings as a JSON array of objects, where each object has the keys "action", "responsible_person", and "deadline". If no responsible person is explicitly mentioned, use "unspecified". If no deadline is mentioned, use null.

Here is the conversation:
{candidate['conversation_text']}"""

        raw_actions_json = self.call_llm(action_prompt)
        
        try:
            extracted_actions = json.loads(raw_actions_json)
        except json.JSONDecodeError:
            extracted_actions = []

        # Generate summary
        summary_prompt = f"""Summarize the following conversation concisely, highlighting the main points and overall outcome.

Conversation:
{candidate['conversation_text']}"""
        
        ai_summary = self.call_llm(summary_prompt)

        # Generate follow-up email
        actions_formatted = "\n".join([
            f"- {item.get('action')} (Responsible: {item.get('responsible_person')}, Due: {item.get('deadline') if item.get('deadline') else 'N/A'})"
            for item in extracted_actions
        ])

        email_generation_prompt = f"""You are an AI assistant tasked with drafting a professional follow-up email based on the following conversation summary and action items.

Conversation Summary:
{ai_summary}

Action Items:
{actions_formatted}

Key Participants: {', '.join(candidate['participants'])}

Draft the follow-up email, including a suitable subject line. Start with "Subject: " for the subject line. Make the email professional and concise. End the email with "Best regards,\nMohammad Reza Kazemi"."""

        generated_email_raw = self.call_llm(email_generation_prompt)

        # Parse subject and body
        subject_match = re.search(r"Subject: (.+?)\n\n(.+)", generated_email_raw, re.DOTALL)
        if subject_match:
            generated_subject = subject_match.group(1).strip()
            generated_body = subject_match.group(2).strip()
        else:
            generated_subject = f"Follow-up on: {candidate['subject']}"
            generated_body = generated_email_raw.strip()
            
        # Ensure signature is included
        if "Mohammad Reza Kazemi" not in generated_body:
            generated_body += "\n\nBest regards,\nMohammad Reza Kazemi"

        return {
            'subject': generated_subject,
            'body': generated_body,
            'recipients': candidate['participants']
        }
        
    def on_email_select(self, event):
        """Handle email selection in listbox"""
        selection = self.emails_listbox.curselection()
        if selection:
            index = selection[0]
            if index < len(self.followup_candidates):
                candidate = self.followup_candidates[index]
                if 'generated_email' in candidate:
                    email_content = candidate['generated_email']
                    preview_text = f"To: {', '.join(email_content['recipients'])}\n"
                    preview_text += f"Subject: {email_content['subject']}\n\n"
                    preview_text += email_content['body']
                    self.email_preview.delete(1.0, tk.END)
                    self.email_preview.insert(1.0, preview_text)
                    # Store selected index for chat revision
                    self.selected_email_index = index
                else:
                    self.selected_email_index = None
            else:
                self.selected_email_index = None
        else:
            self.selected_email_index = None
        
    def revise_email(self):
        """Revise the selected generated email using a user prompt and LLM"""
        prompt = self.chat_var.get().strip()
        if not hasattr(self, 'selected_email_index') or self.selected_email_index is None:
            messagebox.showwarning("Warning", "Please select a follow-up email to revise.")
            return
        if not prompt:
            messagebox.showwarning("Warning", "Please enter a prompt to revise the email.")
            return
        candidate = self.followup_candidates[self.selected_email_index]
        if 'generated_email' not in candidate:
            messagebox.showerror("Error", "No generated email to revise.")
            return
        email_content = candidate['generated_email']
        # Compose revision prompt
        revision_prompt = f"You are an AI assistant. Here is a follow-up email draft:\n\nSubject: {email_content['subject']}\n\n{email_content['body']}\n\nUser request: {prompt}\n\nPlease revise the email accordingly. Start with 'Subject: ' for the subject line."
        self.chat_button.config(state='disabled')
        self.chat_button.update()
        try:
            revised_email_raw = self.call_llm(revision_prompt)
            import re
            subject_match = re.search(r"Subject: (.+?)\n\n(.+)", revised_email_raw, re.DOTALL)
            if subject_match:
                new_subject = subject_match.group(1).strip()
                new_body = subject_match.group(2).strip()
            else:
                new_subject = email_content['subject']
                new_body = revised_email_raw.strip()
            # Ensure signature
            if "Mohammad Reza Kazemi" not in new_body:
                new_body += "\n\nBest regards,\nMohammad Reza Kazemi"
            # Update candidate and preview
            candidate['generated_email'] = {
                'subject': new_subject,
                'body': new_body,
                'recipients': email_content['recipients']
            }
            preview_text = f"To: {', '.join(email_content['recipients'])}\nSubject: {new_subject}\n\n{new_body}"
            self.email_preview.delete(1.0, tk.END)
            self.email_preview.insert(1.0, preview_text)
            self.log_message(f"Email revised for: {candidate['subject']}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to revise email: {str(e)}")
        finally:
            self.chat_button.config(state='normal')
            self.chat_var.set("")
        
    def send_followups(self):
        """Send all follow-up emails"""
        if not self.followup_candidates:
            messagebox.showwarning("Warning", "No follow-up emails to send")
            return
            
        result = messagebox.askyesno("Confirm", f"Send {len(self.followup_candidates)} follow-up emails?")
        if not result:
            return
            
        self.progress_bar.start()
        self.send_button.config(state='disabled')
        
        # Start sending in thread
        thread = threading.Thread(target=self.send_emails_thread)
        thread.daemon = True
        thread.start()
        
    def send_emails_thread(self):
        """Thread function for sending emails"""
        sent_count = 0
        failed_count = 0
        
        for i, candidate in enumerate(self.followup_candidates):
            try:
                self.message_queue.put({'type': 'progress', 'text': f'Sending email {i+1}/{len(self.followup_candidates)}...'})
                
                if 'generated_email' in candidate:
                    email_content = candidate['generated_email']
                    
                    # Use EWS instead of SMTP
                    success = self.send_email_ews(
                        email_content['subject'],
                        email_content['body'],
                        email_content['recipients']
                    )
                    
                    if success:
                        sent_count += 1
                        self.message_queue.put({'type': 'log', 'text': f'Successfully sent follow-up for: {candidate["subject"]}', 'level': 'INFO'})
                    else:
                        failed_count += 1
                        self.message_queue.put({'type': 'log', 'text': f'Failed to send follow-up for: {candidate["subject"]}', 'level': 'ERROR'})
                        
            except Exception as e:
                failed_count += 1
                self.message_queue.put({'type': 'log', 'text': f'Error sending follow-up for "{candidate["subject"]}": {str(e)}', 'level': 'ERROR'})
        
        # Complete
        self.message_queue.put({'type': 'progress', 'text': f'Email sending complete - {sent_count} sent, {failed_count} failed'})
        self.message_queue.put({'type': 'log', 'text': f'Email sending completed: {sent_count} successful, {failed_count} failed', 'level': 'INFO'})
        
        # Re-enable button
        self.root.after(0, lambda: [
            self.progress_bar.stop(),
            self.send_button.config(state='normal')
        ])

        
    def handle_error(self, error_message):
        """Handle errors from worker threads"""
        self.is_processing = False
        self.progress_bar.stop()
        self.fetch_button.config(state='normal')
        self.progress_var.set("Error occurred")
        messagebox.showerror("Error", error_message)
        
    def export_logs(self):
        """Export logs to file"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title="Export Logs"
        )
        
        if filename:
            try:
                logs_content = self.logs_text.get(1.0, tk.END)
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(logs_content)
                messagebox.showinfo("Success", f"Logs exported to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export logs: {str(e)}")
                
    # --- Email and LLM functions (adapted from original code) ---
    
    def call_llm(self, prompt_text):
        """Make API call to LLM"""
        if self.LLM_SIMULATION:
            return "This is a simulated response for testing."
            
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.LLM_API_KEY}"
        }

        payload = {
            "model": self.LLM_MODEL_NAME,
            "messages": [
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt_text}
            ],
            "temperature": 0.2,
            "max_tokens": 500
        }

        try:
            response = requests.post(self.LLM_API_ENDPOINT, headers=headers, json=payload, timeout=120)
            response.raise_for_status()
            response_data = response.json()
            if response_data and response_data.get('choices') and response_data['choices'][0].get('message'):
                return response_data['choices'][0]['message']['content'].strip()
            else:
                return "LLM_ERROR"
        except requests.exceptions.RequestException as e:
            return "LLM_ERROR"
            
    def determine_followup_need(self, conversation_text):
        """Determine if conversation needs follow-up using LLM"""
        prompt = f"""Analyze the following email conversation and determine if it requires a follow-up email.

A follow-up is needed if:
- There are action items or commitments mentioned that need tracking
- Questions were asked but not fully answered
- Decisions were made that require confirmation or next steps
- Meeting outcomes or project status updates are needed
- There are pending deliverables or deadlines

Respond with only "YES" or "NO" followed by a brief explanation.

Conversation:
{conversation_text}"""

        response = self.call_llm(prompt)
        
        if response == "LLM_ERROR":
            return False, "Unable to analyze conversation"
        
        needs_followup = response.upper().startswith("YES")
        explanation = response.split("\n", 1)[1] if "\n" in response else response
        
        return needs_followup, explanation
        
    def get_plain_text_body_ews(self, msg):
        """Extract plain text from EWS message"""
        if msg.body:
            if msg.text_body:
                return msg.text_body
            elif msg.body.content_type == 'HTML' and msg.body.text:
                soup = BeautifulSoup(msg.body.text, 'html.parser')
                return soup.get_text(separator=' ', strip=True)
        return ""
        
    def fetch_emails_from_previous_day_ews(self, days_to_check=1):
        """Fetch emails from the previous N days using EWS"""
        try:
            credentials = Credentials(self.user_email, self.user_password)
            account = Account(primary_smtp_address=self.user_email, credentials=credentials, 
                             autodiscover=True, access_type=DELEGATE)

            # Calculate date range
            today = datetime.now().date()
            start_day = today - timedelta(days=days_to_check)
            start_time = datetime.combine(start_day, datetime.min.time())
            end_time = datetime.combine(today, datetime.max.time())

            # Convert to timezone-aware datetime
            tz = pytz.timezone('UTC')
            start_time = tz.localize(start_time)
            end_time = tz.localize(end_time)

            # Fetch emails from inbox
            messages = account.inbox.filter(
                datetime_received__gte=start_time,
                datetime_received__lte=end_time
            ).order_by('datetime_received')

            # Group emails by conversation
            conversations = {}

            for msg in messages:
                sender_email = msg.sender.email_address if msg.sender else 'unknown'
                sender_name_display = msg.sender.name if msg.sender and msg.sender.name else sender_email.split('@')[0].replace('.', ' ').title()

                # Clean subject for grouping
                clean_subject = re.sub(r'^(Re:|Fwd:|RE:|FWD:)\s*', '', msg.subject or "No Subject", flags=re.IGNORECASE).strip()

                email_data = {
                    "message_id": msg.message_id,
                    "subject": msg.subject,
                    "sender": sender_email,
                    "sender_name_display": sender_name_display,
                    "receivedDateTime": msg.datetime_received.isoformat(),
                    "content": self.get_plain_text_body_ews(msg)
                }

                if clean_subject not in conversations:
                    conversations[clean_subject] = []
                conversations[clean_subject].append(email_data)

            return conversations

        except Exception as e:
            raise Exception(f"Failed to fetch emails: {str(e)}")
            
    def send_email_smtp(self, subject, body, to_addresses):
        """Send email using SMTP"""
        try:
            message = MIMEMultipart("alternative")
            message["From"] = self.user_email
            message["To"] = ", ".join(to_addresses)
            message["Subject"] = subject

            # Convert body to HTML format
            html_body = body.replace('\n', '<br>')
            html_part = MIMEText(html_body, "html")
            message.attach(html_part)

            # Send email
            context = ssl.create_default_context()
            context.check_hostname = False
            context.verify_mode = ssl.CERT_NONE

            if self.SMTP_PORT == 587:
                server = smtplib.SMTP(self.SMTP_SERVER, self.SMTP_PORT)
                server.starttls(context=context)
            elif self.SMTP_PORT == 465:
                server = smtplib.SMTP_SSL(self.SMTP_SERVER, self.SMTP_PORT, context=context)
            else:
                return False

            server.login(self.user_email, self.user_password)
            server.sendmail(self.user_email, to_addresses, message.as_string())
            server.quit()

            return True

        except Exception as e:
            self.message_queue.put({'type': 'log', 'text': f'Email sending error: {str(e)}', 'level': 'ERROR'})
            return False


def main():
    """Main function to run the application"""
    root = tk.Tk()
    app = EmailFollowupApp(root)
    
    # Center the window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("Application interrupted by user")
    except Exception as e:
        print(f"Application error: {e}")


if __name__ == "__main__":
    main()