#!/usr/bin/env python3
"""
================================================================================
CDMP QUIZ APPLICATION - CLEAN & OPTIMIZED VERSION
================================================================================
A streamlined quiz application for CDMP exam practice.
Author: Mostafa Maher
Link:https://github.com/FINAL094
www.linkedin.com/in/eng-mostafa-maher

Requirements:
    pip install pandas openpyxl

Usage:
    1. Place this script next to 'CDMP Practice Exam v4.xlsx'
    2. Run: python cdmp_quiz.py
================================================================================
"""

import os
import sys
import time
import re
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

# ============================================================================
# CONFIGURATION SECTION
# ============================================================================

EXCEL_FILE = "CDMP Practice Exam.xlsx"
SECONDS_PER_QUESTION = 30
AUTO_ADVANCE_DELAY = 700

CHAPTERS = {
    "1": "Data Management",
    "2": "Data Governance",
    "3": "Data Handling Ethics",
    "4": "Data Quality",
    "5": "Data Management Organization and Role Expectation",
    "6": "Reference & Master Data",
    "7": "Big Data and Data Science",
    "8": "Data Security",
    "9": "Data Architecture",
    "10": "Data Integration & Interoperability",
    "11": "Data Modeling and Design",
    "12": "Metadata",
    "13": "Data Warehousing and Business Intelligence",
    "14": "Data Storage and Operations",
    "15": "Document and Content Management",
}

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def set_working_directory():
    """Set working directory to script location."""
    try:
        script_path = os.path.abspath(sys.argv[0])
        directory = os.path.dirname(script_path)
        if directory:
            os.chdir(directory)
    except:
        pass

def normalize_chapter_name(value):
    """Convert chapter identifiers to standard names."""
    if pd.isna(value) or str(value).strip() == "":
        return "Unspecified"
    
    try:
        num = float(value)
        if num.is_integer():
            key = str(int(num))
            if key in CHAPTERS:
                return CHAPTERS[key]
    except:
        pass
    
    if value in CHAPTERS.values():
        return value
    
    return str(value).strip()

def extract_correct_letters(text):
    """Extract correct answer letters (A-E) from text."""
    if pd.isna(text):
        return []
    
    items = re.split(r'[,;/\s]+', str(text).strip())
    letters = []
    
    for item in items:
        match = re.match(r'([A-Ea-e])', item)
        if match:
            letters.append(match.group(1).upper())
    
    return list(dict.fromkeys(letters))

# ============================================================================
# MAIN APPLICATION CLASS
# ============================================================================

class CDMPQuizApp:
    """Main quiz application class."""
    
    def __init__(self, master, questions_df, answers_df):
        """Initialize the quiz application."""
        self.master = master
        self.master.title("CDMP Quiz Application")
        self.master.geometry("980x720")
        
        self.all_questions = questions_df.copy().reset_index(drop=True)
        self.all_answers = answers_df.copy().reset_index(drop=True)
        
        self.questions = pd.DataFrame()
        self.answers = pd.DataFrame()
        
        self.active = False
        self.review_mode = False
        self.timer_active = False
        self.timer_expired = False
        self.start_time = None
        self.total_time = 0
        
        self.current_index = 0
        self.current_question_idx = None
        self.question_start_time = None
        
        self.user_answers = []
        
        self.answer_widgets = []
        self.selected_radio = tk.StringVar(value='NONE')
        self.selected_checkboxes = []
        
        self.create_interface()
        self.show_greeting()
    
    def show_greeting(self):
        """Display Islamic greeting popup."""
        greeting = """بسم الله الرحمن الرحيم

صلى الله على النبي محمد عليه أفضل الصلاة والسلام

وادعو لي لو استفدت من البرنامج"""
        messagebox.showinfo("بركة", greeting)
    
    def create_interface(self):
        """Build the main user interface."""
        self.start_frame = tk.Frame(self.master, padx=10, pady=8)
        self.quiz_frame = tk.Frame(self.master, padx=10, pady=8)
        
        self.create_start_screen()
        self.create_quiz_screen()
        
        self.start_frame.pack(fill='both', expand=True)
    
    def create_start_screen(self):
        """Create the start/welcome screen."""
        greeting_box = tk.Frame(self.start_frame, bg="#f0f8ff", relief="ridge", borderwidth=2)
        greeting_box.grid(row=0, column=0, columnspan=3, sticky='ew', pady=(0, 15))
        
        greeting_text = ("بسم الله الرحمن الرحيم\n"
                        "صلى الله على النبي محمد عليه أفضل الصلاة والسلام\n"
                        "وادعو لي لو استفدت من البرنامج")
        
        tk.Label(greeting_box, text=greeting_text, font=("Arial", 11),
                bg="#f0f8ff", fg="#1a5490", justify='center', pady=10).pack()
        
        tk.Label(self.start_frame, text="CDMP Quiz Application",
                font=("Helvetica", 18, "bold")).grid(row=1, column=0, columnspan=3, sticky='w', pady=(0, 12))
        
        tk.Label(self.start_frame, text="Select Chapter:",
                font=("Helvetica", 12)).grid(row=2, column=0, sticky='w')
        
        self.chapter_var = tk.StringVar()
        chapters = self.get_chapter_list()
        self.chapter_combo = ttk.Combobox(self.start_frame, values=chapters,
                                         textvariable=self.chapter_var,
                                         state='readonly', width=55)
        self.chapter_combo.grid(row=2, column=1, sticky='w', padx=(8, 8))
        self.chapter_combo.set(chapters[0])
        self.chapter_combo.bind("<<ComboboxSelected>>", lambda e: self.update_chapter_info())
        
        self.randomize_questions = tk.IntVar(value=0)
        self.randomize_answers = tk.IntVar(value=0)
        
        tk.Checkbutton(self.start_frame, text="Randomize Question Order",
                      variable=self.randomize_questions).grid(row=3, column=0, columnspan=3, sticky='w', pady=(8, 0))
        
        tk.Checkbutton(self.start_frame, text="Randomize Answer Order",
                      variable=self.randomize_answers).grid(row=4, column=0, columnspan=3, sticky='w')
        
        tk.Button(self.start_frame, text="Start Quiz", width=16, bg="#4caf50", fg="white",
                 command=self.start_quiz).grid(row=2, column=2, rowspan=2, padx=(6, 0))
        
        self.info_label = tk.Label(self.start_frame, text="", fg="gray", font=("Helvetica", 10))
        self.info_label.grid(row=5, column=0, columnspan=3, pady=(12, 0), sticky='w')
        self.update_chapter_info()
    
    def create_quiz_screen(self):
        """Create the quiz/question screen."""
        topbar = tk.Frame(self.quiz_frame)
        topbar.pack(fill='x', pady=(4, 6))
        
        self.timer_label = tk.Label(topbar, text="Time: 00:00", font=("Helvetica", 12))
        self.timer_label.pack(side='left')
        
        tk.Button(topbar, text="End Quiz", bg="#d9534f", fg="white",
                 command=self.end_quiz).pack(side='right', padx=4)
        
        tk.Button(topbar, text="← Back to Start",
                 command=self.back_to_start).pack(side='right', padx=4)
        
        self.show_answer_btn = tk.Button(topbar, text="Show Answer",
                                         command=self.reveal_answer, state='disabled')
        self.show_answer_btn.pack(side='right', padx=4)
        
        self.review_btn = tk.Button(topbar, text="Review Exam",
                                    command=self.start_review, state='disabled')
        self.review_btn.pack(side='right', padx=4)
        
        self.question_label = tk.Label(self.quiz_frame, text="Question will appear here",
                                      wraplength=920, font=("Helvetica", 17),
                                      justify='left', anchor='w')
        self.question_label.pack(fill='x', pady=(6, 8))
        
        self.answers_container = tk.Frame(self.quiz_frame)
        self.answers_container.pack(fill='both', expand=True, anchor='nw')
        
        nav_frame = tk.Frame(self.quiz_frame)
        nav_frame.pack(fill='x', pady=(8, 6))
        
        self.prev_btn = tk.Button(nav_frame, text="← Previous", width=14,
                                  state='disabled', command=self.go_previous)
        self.prev_btn.pack(side='left', padx=6)
        
        self.submit_btn = tk.Button(nav_frame, text="Submit Answer", width=14,
                                    bg="#1976d2", fg="white", state='disabled',
                                    command=self.submit_answer)
        self.submit_btn.pack(side='left', padx=6)
        
        self.skip_btn = tk.Button(nav_frame, text="Skip →", width=14,
                                  state='disabled', command=self.skip_question)
        self.skip_btn.pack(side='left', padx=6)
        
        self.status_label = tk.Label(self.quiz_frame, text="", font=("Helvetica", 11))
        self.status_label.pack(anchor='w', pady=(6, 0))
        
        self.master.bind("<Key>", self.handle_keypress)
    
    def get_chapter_list(self):
        """Get list of chapters for dropdown."""
        chapters = ["All Chapters"]
        for key in sorted(CHAPTERS.keys(), key=int):
            chapters.append(f"{key} - {CHAPTERS[key]}")
        return chapters
    
    def get_selected_chapter(self):
        """Get the currently selected chapter name."""
        selection = self.chapter_var.get().strip()
        if not selection or selection == "All Chapters":
            return "All Chapters"
        if " - " in selection:
            return selection.split(" - ", 1)[1].strip()
        return selection
    
    def update_chapter_info(self):
        """Update the chapter info display."""
        chapter = self.get_selected_chapter()
        
        if chapter == "All Chapters":
            count = len(self.all_questions)
            display = "All Chapters"
        else:
            count = len(self.all_questions[self.all_questions['chapter'] == chapter])
            display = chapter
        
        self.info_label.config(text=f"Chapter: {display}   |   Questions: {count}")
    
    def start_quiz(self):
        """Initialize and start the quiz."""
        chapter = self.get_selected_chapter()
        
        if chapter == "All Chapters":
            questions = self.all_questions.copy().reset_index(drop=True)
        else:
            questions = self.all_questions[self.all_questions['chapter'] == chapter].copy().reset_index(drop=True)
        
        if questions.empty:
            messagebox.showwarning("No Questions", f"No questions found for: {chapter}")
            return
        
        if self.randomize_questions.get() == 1:
            questions = questions.sample(frac=1).reset_index(drop=True)
        
        question_ids = questions['qid'].unique()
        answer_frames = []
        
        for qid in question_ids:
            qanswers = self.all_answers[self.all_answers['qid'] == qid].copy().reset_index(drop=True)
            
            if self.randomize_answers.get() == 1:
                qanswers = qanswers.sample(frac=1).reset_index(drop=True)
            
            answer_frames.append(qanswers)
        
        answers = pd.concat(answer_frames, ignore_index=True) if answer_frames else pd.DataFrame()
        
        self.questions = questions
        self.answers = answers
        
        self.user_answers = []
        self.current_index = 0
        self.current_question_idx = None
        self.active = True
        self.review_mode = False
        self.timer_active = False
        self.timer_expired = False
        self.start_time = None
        self.total_time = len(questions) * SECONDS_PER_QUESTION
        
        self.start_frame.pack_forget()
        self.quiz_frame.pack(fill='both', expand=True)
        self.status_label.config(text=f"Chapter: {chapter}   |   Questions: {len(questions)}")
        
        self.load_next_question()
    
    def load_next_question(self):
        """Load and display the next question."""
        if not self.active:
            return
        
        self.clear_answer_widgets()
        self.selected_radio.set('NONE')
        self.selected_checkboxes = []
        self.submit_btn.config(state='disabled')
        self.skip_btn.config(state='disabled')
        self.show_answer_btn.config(state='disabled')
        
        if self.current_index < len(self.questions):
            question = self.questions.iloc[self.current_index]
            self.current_question_idx = self.current_index
            self.current_index += 1
            
            chapter = self.get_selected_chapter()
            qid = question.get('qid', '')
            header = f"Chapter: {chapter}   |   QID: {qid}   ({self.current_question_idx + 1} of {len(self.questions)})\n\n"
            self.question_label.config(text=header + str(question.get('question', '')))
            
            self.current_metadata = {'qid': qid, 'type': question.get('type', 'single')}
            self.question_start_time = time.time()
            
            if not self.timer_active:
                self.start_time = time.time()
                self.update_timer()
                self.timer_active = True
            
            self.display_answers(qid, question.get('type', 'single'))
            
            self.skip_btn.config(state='normal')
            self.prev_btn.config(state='normal' if self.current_question_idx > 0 else 'disabled')
        else:
            self.finish_quiz()
    
    def go_previous(self):
        """Go back to the previous question."""
        if self.current_question_idx is None or self.current_question_idx == 0:
            return
        
        # Remove answer for current question if exists
        current_qid = self.questions.iloc[self.current_question_idx]['qid']
        self.user_answers = [ans for ans in self.user_answers if ans.get('qid') != current_qid]
        
        # Go back one question
        self.current_index = self.current_question_idx - 1
        self.load_next_question()
    
    def display_answers(self, qid, question_type, show_correct=False, user_response=None):
        """Display answer options for current question."""
        self.clear_answer_widgets()
        
        qanswers = self.answers[self.answers['qid'] == qid].reset_index(drop=True)
        
        if qanswers.empty:
            label = tk.Label(self.answers_container, text="(No answers available)",
                           font=("Helvetica", 14))
            label.pack(anchor='w', padx=8, pady=6)
            self.answer_widgets.append(label)
            return
        
        for idx, answer in qanswers.iterrows():
            option_text = str(answer.get('options', '')).strip()
            reference = str(answer.get('ref', '')).strip()
            value = answer.get('value', '')
            is_correct = int(answer.get('point', 0)) == 1
            
            # No numbering - just option text
            display = option_text
            
            if show_correct or self.review_mode or self.timer_expired:
                if is_correct:
                    display = f"{option_text}  ✓ (Correct)"
                elif user_response is not None:
                    if question_type == 'single' and str(user_response) == str(value):
                        display = f"{option_text}  ✘ (Your Answer)"
                    elif (question_type != 'single' and isinstance(user_response, (list, tuple)) and
                          str(value) in [str(x) for x in user_response]):
                        display = f"{option_text}  ✘ (Your Answer)"
                
                if reference:
                    display += f"  → {reference}"
            
            state = 'disabled' if (show_correct or self.review_mode or self.timer_expired) else 'normal'
            bg_color = None
            
            if show_correct or self.review_mode or self.timer_expired:
                if is_correct:
                    bg_color = 'lightgreen'
                elif user_response is not None:
                    if question_type == 'single' and str(user_response) == str(value):
                        bg_color = 'lightcoral'
                    elif (question_type != 'single' and isinstance(user_response, (list, tuple)) and
                          str(value) in [str(x) for x in user_response]):
                        bg_color = 'lightcoral'
            
            if question_type == 'single':
                widget = tk.Radiobutton(self.answers_container, text=display,
                                       variable=self.selected_radio, value=value,
                                       command=self.on_answer_selected, anchor='w',
                                       justify='left', state=state, wraplength=900,
                                       font=("Helvetica", 14))
            else:
                var = tk.StringVar(value='')
                widget = tk.Checkbutton(self.answers_container, text=display,
                                       variable=var, onvalue=value, offvalue='',
                                       command=lambda v=var: self.on_answer_selected(v),
                                       anchor='w', justify='left', state=state,
                                       wraplength=900, font=("Helvetica", 14))
                self.selected_checkboxes.append(var)
            
            widget._value = value
            widget._is_correct = is_correct
            widget._option_text = option_text
            widget._reference = reference
            widget._index = idx
            
            if bg_color:
                try:
                    widget.configure(bg=bg_color)
                except:
                    pass
            
            widget.pack(anchor='w', padx=8, pady=4)
            self.answer_widgets.append(widget)
    
    def clear_answer_widgets(self):
        """Remove all answer widgets from display."""
        for widget in self.answer_widgets:
            try:
                widget.destroy()
            except:
                pass
        self.answer_widgets = []
    
    def on_answer_selected(self, checkbox_var=None):
        """Handle answer selection (radio or checkbox)."""
        self.submit_btn.config(state='normal')
        
        if not self.review_mode:
            self.show_answer_btn.config(state='normal')
        
        if checkbox_var is not None:
            selected = [v.get() for v in self.selected_checkboxes if v.get() != '']
            if not selected:
                self.submit_btn.config(state='disabled')
                self.show_answer_btn.config(state='disabled')
    
    def submit_answer(self):
        """Submit the current answer and move to next question."""
        if self.current_question_idx is None or not self.active:
            return
        
        elapsed = int(time.time() - self.question_start_time) if self.question_start_time else 0
        
        qid = self.current_metadata.get('qid')
        qtype = self.current_metadata.get('type', 'single')
        
        if qtype == 'single':
            user_answer = self.selected_radio.get()
            if user_answer == 'NONE':
                user_answer = ''
        else:
            user_answer = [v.get() for v in self.selected_checkboxes if v.get() != '']
        
        points = self.calculate_points(qid, user_answer)
        
        self.user_answers.append({
            'qid': qid,
            'type': qtype,
            'answer': user_answer,
            'points': points,
            'time': elapsed
        })
        
        self.update_answer_display(qid, user_answer)
        
        self.submit_btn.config(state='disabled')
        self.skip_btn.config(state='disabled')
        self.show_answer_btn.config(state='disabled')
        
        self.master.after(AUTO_ADVANCE_DELAY, lambda: self.load_next_question() if self.active else None)
    
    def skip_question(self):
        """Skip current question without recording an answer."""
        self.load_next_question()
    
    def update_answer_display(self, qid, user_answer):
        """Update answer widgets to show correct/incorrect."""
        for widget in self.answer_widgets:
            value = getattr(widget, '_value', None)
            is_correct = getattr(widget, '_is_correct', False)
            option_text = getattr(widget, '_option_text', '')
            reference = getattr(widget, '_reference', '')
            
            display = option_text
            
            if is_correct:
                display = f"{option_text}  ✓ (Correct)"
            elif user_answer is not None:
                if isinstance(user_answer, list):
                    if str(value) in [str(x) for x in user_answer]:
                        display = f"{option_text}  ✘ (Your Answer)"
                else:
                    if str(user_answer) == str(value):
                        display = f"{option_text}  ✘ (Your Answer)"
            
            if reference:
                display += f"  → {reference}"
            
            bg_color = None
            if is_correct:
                bg_color = 'lightgreen'
            elif user_answer is not None:
                if isinstance(user_answer, list):
                    if str(value) in [str(x) for x in user_answer]:
                        bg_color = 'lightcoral'
                elif str(user_answer) == str(value):
                    bg_color = 'lightcoral'
            
            try:
                widget.configure(text=display, state='disabled')
                if bg_color:
                    widget.configure(bg=bg_color)
            except:
                pass
    
    def calculate_points(self, qid, user_selection):
        """Calculate points earned for an answer."""
        correct = self.answers[(self.answers['qid'] == qid) & (self.answers['point'] == 1)]
        
        if correct.empty:
            return 0.0
        
        points_per_correct = 1.0 / len(correct)
        
        if not user_selection or user_selection == '' or user_selection == 'NONE':
            return 0.0
        
        if isinstance(user_selection, list):
            correct_values = set(correct['value'].astype(str))
            selected_values = set(str(x) for x in user_selection)
            matches = len(correct_values & selected_values)
            return float(matches * points_per_correct)
        else:
            correct_values = set(correct['value'].astype(str))
            return 1.0 if str(user_selection) in correct_values else 0.0
    
    def reveal_answer(self):
        """Show correct answers for current question."""
        if self.current_question_idx is None:
            return
        
        qid = self.questions.iloc[self.current_question_idx]['qid']
        
        user_answer = None
        for record in reversed(self.user_answers):
            if record.get('qid') == qid:
                user_answer = record.get('answer')
                break
        
        self.update_answer_display(qid, user_answer)
        self.submit_btn.config(state='disabled')
    
    def update_timer(self):
        """Update the timer display."""
        if not self.start_time:
            self.start_time = time.time()
        
        elapsed = int(time.time() - self.start_time)
        remaining = max(0, self.total_time - elapsed)
        
        minutes, seconds = divmod(remaining, 60)
        self.timer_label.config(text=f"Time: {minutes:02d}:{seconds:02d}")
        
        if remaining > 0 and not self.timer_expired and self.active:
            self.master.after(1000, self.update_timer)
        else:
            if not self.timer_expired:
                self.timer_expired = True
                if self.active:
                    messagebox.showinfo("Time's Up", "Timer expired. Entering review mode.")
                    
                    if (self.current_question_idx is not None and
                        self.current_question_idx < len(self.questions)):
                        qid = self.questions.iloc[self.current_question_idx]['qid']
                        user_answer = None
                        for record in reversed(self.user_answers):
                            if record.get('qid') == qid:
                                user_answer = record.get('answer')
                                break
                        self.update_answer_display(qid, user_answer)
                    
                    self.active = False
                    self.review_btn.config(state='normal')
                    self.submit_btn.config(state='disabled')
                    self.skip_btn.config(state='disabled')
                    self.show_answer_btn.config(state='disabled')
    
    def end_quiz(self):
        """End quiz early (with confirmation)."""
        if not messagebox.askyesno("End Quiz", "End quiz now and view results?"):
            return
        self.active = False
        self.show_results_and_reset()
    
    def finish_quiz(self):
        """Quiz finished naturally (all questions completed)."""
        self.active = False
        self.show_results_and_reset()
    
    def show_results_and_reset(self):
        """Display final results and reset to start screen."""
        if self.user_answers:
            results_df = pd.DataFrame(self.user_answers)
            total_points = float(results_df['points'].sum())
            attempted = len(results_df)
        else:
            total_points = 0.0
            attempted = 0
        
        total_questions = len(self.questions)
        chapter = self.get_selected_chapter()
        
        messagebox.showinfo("Quiz Complete",
                          f"Chapter: {chapter}\n"
                          f"Score: {total_points:.1f} / {total_questions}\n"
                          f"Attempted: {attempted}")
        
        self.show_greeting()
        
        self.reset_quiz_state()
        self.show_start_screen()
    
    def reset_quiz_state(self):
        """Reset all quiz state variables."""
        self.user_answers = []
        self.current_index = 0
        self.current_question_idx = None
        self.question_start_time = None
        self.start_time = None
        self.timer_active = False
        self.timer_expired = False
        self.active = False
        self.review_mode = False
        self.questions = pd.DataFrame()
        self.answers = pd.DataFrame()
        self.clear_answer_widgets()
        self.question_label.config(text="Question will appear here")
        self.timer_label.config(text="Time: 00:00")
        self.status_label.config(text="")
    
    def show_start_screen(self):
        """Switch back to start screen."""
        self.quiz_frame.pack_forget()
        self.start_frame.pack(fill='both', expand=True)
        self.update_chapter_info()
        self.chapter_combo['values'] = self.get_chapter_list()
        if self.chapter_var.get() not in self.chapter_combo['values']:
            self.chapter_combo.set(self.chapter_combo['values'][0])
    
    def back_to_start(self):
        """Return to start screen (with confirmation if quiz active)."""
        if self.active:
            if not messagebox.askyesno("Confirm", "Go back to start? This will reset progress."):
                return
        self.active = False
        self.reset_quiz_state()
        self.show_start_screen()
    
    def start_review(self):
        """Enter review mode to see all questions and answers."""
        if not self.user_answers:
            messagebox.showinfo("No Answers", "No recorded answers to review.")
            return
        
        self.review_mode = True
        self.review_index = 0
        self.load_review_question()
    
    def load_review_question(self):
        """Load next question in review mode."""
        self.clear_answer_widgets()
        
        if self.review_index < len(self.questions):
            question = self.questions.iloc[self.review_index]
            qid = question.get('qid')
            qtype = question.get('type', 'single')
            
            chapter = self.get_selected_chapter()
            header = f"Chapter: {chapter}   |   QID: {qid}   ({self.review_index + 1} of {len(self.questions)})\n\n"
            self.question_label.config(text=header + str(question.get('question', '')))
            
            user_answer = None
            for record in self.user_answers:
                if record.get('qid') == qid:
                    user_answer = record.get('answer')
                    break
            
            self.display_answers(qid, qtype, show_correct=True, user_response=user_answer)
            self.review_index += 1
        else:
            messagebox.showinfo("Review Complete", "End of review.")
            self.review_mode = False
    
    def handle_keypress(self, event):
        """Handle keyboard shortcuts (1-9 for answer selection)."""
        key = event.char.strip()
        if not key.isdigit():
            return
        
        idx = int(key) - 1
        if idx < 0 or idx >= len(self.answer_widgets):
            return
        
        try:
            self.answer_widgets[idx].invoke()
        except:
            pass

# ============================================================================
# DATA LOADING FUNCTIONS
# ============================================================================

def load_excel_data(filepath):
    """Load and parse Excel file into questions and answers dataframes."""
    try:
        xls = pd.ExcelFile(filepath, engine='openpyxl')
    except Exception as e:
        raise Exception(f"Error opening Excel file: {e}")
    
    sheet_names_lower = [s.lower() for s in xls.sheet_names]
    
    if 'ques' in sheet_names_lower and 'ans' in sheet_names_lower:
        ques_sheet = [s for s in xls.sheet_names if s.lower() == 'ques'][0]
        ans_sheet = [s for s in xls.sheet_names if s.lower() == 'ans'][0]
        
        questions = pd.read_excel(xls, sheet_name=ques_sheet, engine='openpyxl')
        answers = pd.read_excel(xls, sheet_name=ans_sheet, engine='openpyxl')
    else:
        first_sheet = xls.sheet_names[0]
        raw_data = pd.read_excel(xls, sheet_name=first_sheet, engine='openpyxl')
        
        questions, answers = parse_combined_format(raw_data)
    
    questions = normalize_questions(questions)
    answers = normalize_answers(answers)
    
    return questions, answers

def parse_combined_format(raw_df):
    """Parse combined format Excel (single sheet with all data)."""
    cols = {c.strip().lower(): c for c in raw_df.columns}
    
    def find_column(variants):
        for var in variants:
            if var.lower() in cols:
                return cols[var.lower()]
        return None
    
    col_chapter = find_column(['KnowledgeArea', 'Knowledge Area', 'knowledgearea'])
    col_qnum = find_column(['Question Number', 'QuestionNumber', 'QNumber'])
    col_question = find_column(['Question', 'question'])
    col_correct = find_column(['Correct', 'correct', 'Answer'])
    col_section = find_column(['DMBOK Section', 'DMBOKSection'])
    col_page = find_column(['DMBOK Page', 'DMBOKPage'])
    
    col_options = {}
    for letter in ['A', 'B', 'C', 'D', 'E']:
        col_options[letter] = find_column([letter, f'{letter}.', letter.lower()])
    
    questions_list = []
    answers_list = []
    
    for idx, row in raw_df.iterrows():
        if col_qnum and pd.notna(row.get(col_qnum)):
            qid = str(row.get(col_qnum)).strip()
        else:
            qid = f"Q{idx + 1}"
        
        if col_chapter and pd.notna(row.get(col_chapter)):
            chapter = normalize_chapter_name(row.get(col_chapter))
        elif col_section and pd.notna(row.get(col_section)):
            chapter = normalize_chapter_name(row.get(col_section))
        else:
            chapter = "Unspecified"
        
        question_text = row.get(col_question) if col_question else ""
        if pd.isna(question_text):
            question_text = ""
        
        correct_letters = extract_correct_letters(row.get(col_correct)) if col_correct else []
        question_type = 'single' if len(correct_letters) <= 1 else 'multiple'
        
        ref_parts = []
        if col_section and pd.notna(row.get(col_section)):
            ref_parts.append(str(row.get(col_section)).strip())
        if col_page and pd.notna(row.get(col_page)):
            ref_parts.append(str(row.get(col_page)).strip())
        reference = " | ".join(ref_parts) if ref_parts else ""
        
        questions_list.append({
            'qid': qid,
            'chapter': chapter,
            'question': str(question_text),
            'type': question_type
        })
        
        for letter in ['A', 'B', 'C', 'D', 'E']:
            col_name = col_options.get(letter)
            if not col_name:
                continue
            
            option_text = row.get(col_name)
            if pd.isna(option_text) or str(option_text).strip() == "":
                continue
            
            is_correct = 1 if letter in correct_letters else 0
            
            answers_list.append({
                'qid': qid,
                'options': str(option_text),
                'value': letter,
                'point': is_correct,
                'randomize': 1,
                'ref': reference
            })
    
    return pd.DataFrame(questions_list), pd.DataFrame(answers_list)

def normalize_questions(df):
    """Ensure questions dataframe has required columns."""
    if 'chapter' not in df.columns:
        df['chapter'] = ''
    df['chapter'] = df['chapter'].fillna('').astype(str).apply(normalize_chapter_name)
    
    if 'type' not in df.columns:
        df['type'] = 'single'
    df['type'] = df['type'].fillna('single').astype(str).str.strip()
    
    if 'qid' not in df.columns:
        df['qid'] = [f"Q{i+1}" for i in range(len(df))]
    
    return df

def normalize_answers(df):
    """Ensure answers dataframe has required columns."""
    for col in ['qid', 'options', 'value', 'point', 'randomize', 'ref']:
        if col not in df.columns:
            df[col] = None
    
    try:
        df['point'] = df['point'].fillna(0).astype(int)
    except:
        df['point'] = df['point'].fillna(0).apply(
            lambda x: int(float(x)) if pd.notna(x) and str(x).replace('.', '', 1).isdigit() else 0
        )
    
    return df

# ============================================================================
# MAIN FUNCTION
# ============================================================================

def main():
    """Main entry point for the application."""
    set_working_directory()
    
    if not os.path.exists(EXCEL_FILE):
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("File Not Found",
                           f"Could not find '{EXCEL_FILE}' in {os.getcwd()}")
        return
    
    try:
        questions, answers = load_excel_data(EXCEL_FILE)
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Load Error", str(e))
        return
    
    root = tk.Tk()
    app = CDMPQuizApp(root, questions, answers)
    root.mainloop()

if __name__ == "__main__":
    main()