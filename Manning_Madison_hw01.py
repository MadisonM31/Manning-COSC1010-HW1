# Madison Manning           
# UWYO COSC 1010
# 10/15/2024
# HW 01
# Lab Section: 10
# Sources, people worked with, help given to: Thank you to Iqbal Khatoon during office hours and my COSC1010 lecture notes

# Homework Question:
# 
# You are given a list of dictionaries where each dictionary represents a student and their scores 
# in different subjects.
# 
# Student Data:
students = [
     {"name": "Alice", "scores": {"Math": 85, "Science": 90, "English": 78}},
     {"name": "Bob", "scores": {"Math": 70, "Science": 88, "English": 82}},
     {"name": "Charlie", "scores": {"Math": 92, "Science": 81, "English": 89}},
     {"name": "David", "scores": {"Math": 60, "Science": 75, "English": 80}}
 ]

#Write a Python program that:
# 1. Calculates the average score for each student.
# 2. Stores these averages in a new dictionary where the student’s name is the key and their average score is the value.
# 3. Prints the names of students whose average score is greater than 80.

# Your task is to calculate the average scores for each student and print the names of students
# whose average score is greater than 80.

#Solution

student_avrg = {}

for student in students: 
    student_avrg[student["name"]] = sum(student["scores"].values()) / 3

print(student_avrg)

for name, score in student_avrg.items():
    if score > 80:
        print(f"{name} has an average score above 80")
    
    

   
    

         

        



