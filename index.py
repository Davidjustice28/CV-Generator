import aspose.words as asp
import datetime
import time
import sys

current_time = datetime.datetime.now()
weekdays = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
months = ("Jan","Feb","Mar","Apr","May","Jun","July","Aug","Sep","Oct","Nov","Dec")
todays_date = f"{weekdays[current_time.weekday()]}, {months[current_time.month]} {current_time.day} {current_time.year}"

def create_new_cover_letter():
    print("Welcome to the CV Generator Program")
    first_name = input("Please enter your first name: ")
    last_name = input("Last Name: ")
    title = input("What is your current job title? ")
    company = input("what is the name of the company offering the job? ")
    position = input("What position are you applying for? ")
    have_related_experience = input("Do you have any related experience to this role? ")
    experiences = []
    technologies = []

    while have_related_experience.lower() == 'yes':
        related_experience = input("What experience do you have in short answer? Ex: building web apps, customer service, etc. ")
        experiences.append(related_experience)
        more_experiences = input("Any other related experiences? ")

        if more_experiences.lower() == 'no':
            have_related_experience = 'no'
            break
    while True:
        knows_a_technology = input("Do you know any industry desired technologies? ")
        if knows_a_technology.lower() == 'yes':
            technologies.append(input("What technology? "))
        else:
            break

    experiences_sentence = ""
    tech_sentence = ""

    for index, e in enumerate(experiences):
        if index == (len(experiences) - 1):
            experiences_sentence += f" and {e}"
        else:
            experiences_sentence += f"{e}, "

    for index, tech in enumerate(technologies):
        if index == (len(technologies) - 1):
            tech_sentence += f" and {tech}"
        else:
            tech_sentence += f"{tech}, "


    print("Sounds like an awesome job!\nPlease wait while I make your cover letter...")
    new_doc = asp.Document()
    documentBuilder = asp.DocumentBuilder(new_doc)

    paragraph_format = documentBuilder.paragraph_format
    paragraph_format.alignment = asp.ParagraphAlignment.LEFT
    paragraph_format.first_line_indent = 8

    documentBuilder.writeln(todays_date)
    documentBuilder.writeln("To whom this may concern,")
    documentBuilder.writeln(f"My name is {first_name+' '+last_name} and I am reaching out, because I am interested in your \
                            position as a {position} for {company}. I am a {title} that specializes in {tech_sentence}. I believe that I would be a great fit for the position.\
                            I have plenty of related experience including {experiences_sentence}. I also work great in team settings\
                            and able to utilize different industry standard tools including version control, Github, Jira, and Slack.\
                            I have a strong a passion for building complex and interactive application that improves the lives\
                            of others and truly enjoy learning new technologies. I would love to setup a time to discuss the opening\
                            and further go over my experience and how I could be a good fit for the role.")
    documentBuilder.writeln("Thank you for taking the time to review my application. I have attached a copy of my resume below.\
                            I look forward to speaking with you soon,")
    documentBuilder.writeln(f"{first_name+' '+last_name}")
    file_name = input("What would you like to call this file? ")
    new_doc.save(file_name+".docx")
    print("Okay your cover letter has been created and saved!")

create_new_cover_letter()