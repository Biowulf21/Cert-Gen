# Cert-Gen

Cert-Gen is a cross-platform, Python-based tool that helps in generating certificates ~~and other documents that needs the names of participants~~ and automatically sends them to their respective emails.

## Table of Contents

- [Cert-Gen](#cert-gen)
  - [Table of Contents](#table-of-contents)
  - [Requirements](#requirements)
  - [Installation](#installation)
  - [Use-Cases](#use-cases)
  - [Limitations](#limitations)
  - [Usage](#usage)
- [FAQs](#faqs) - [This software was created by James Jilhaney](#this-software-was-created-by-james-jilhaney)

## Requirements

- Works with Windows, Linux and Mac
- [Python](https://www.python.org/downloads/) 3.0 or greater
- [Git](https://git-scm.com/downloads)

## Installation

1. Go to the folder where you want to store Cert-Gen.
1. Open the commandline from the chosen folder.
1. Write the following command on the commandline:<br/> `git clone https://github.com/Biowulf21/Cert-Gen.git`
1. type `python3 install -r requirements.txt`
1. Voilà, you know have a copy of the program!

## Use-Cases

Cert-Gen is primarily used in the following use-cases:

- Certificate Generation and Sending
- Receipt Generation and Sending

**Note: Cert-Gen can be used in conjunction with any task that requires the names of multiple people on the same document and for that document to be sent to their respective emails.**

## Limitations

- As of the moment, Cert-Gen does not have a way to confirm if the email has been successfully sent. However, this can be checked in the email sentbox.
- At this point in time, a graphical user interface **(GUI)** has **NOT** been designed (as this script was just written as a quick way to send certificates to emails)

- I will be working on making Cert-Gen an excecutable file ~~as I cannot, for the life of me, figure out how to make that work with Python~~ for easier installation.

## Usage

1. Download all necessary files and dependencies from steps above.
1. Input the names of the receipients on column A and their emails on column B. Make sure the emails are aligned respectively.
1. Make sure to **save** the names.xlsx file. This is important as the program will only detect the values inputted in the worksheet when it is saved.
   **Note:** The data in the sheet should have no blank cells in between. Though this won't cause any crashes, the program will use a None value to create a certificate and send it to an 'None' email.
1. Open the commandline from the folder where Cert-Gen is located at.
1. Create a `Certificates` folder in the root of the project. This is where the generated files will be saved.
1. Create a `.env` file in the root of the project using the following syntax:

```
EMAIL=foo
PASSWORD=foo
```

**Note:** Make sure to replace "foo" with the values you'll be using for the email and password. Also make sure that the values are not enclosed in quotation marks.

Make sure to login to your google account. Go to the security tab and turn the `Less Secure Apps Access` **on**. For more information access this [link](https://support.google.com/accounts/answer/6010255?hl=en).

### On Windows:

    1. Open the folder where Cert-Gen is.
    1. Right click on any blank space in the file manager.
    1. Click on 'Open in terminal'

1. Run the command `python3 certificate.py`
1. You should now see the program generate the certificates and send them to their respective emails :)

# FAQs

- **Q:** Does this program have a Graphical User Interface? <br/>
  **A:** No, a GUI is definitely planned, but right now it's commandline only. It's very easy to use, though.

- **Q:** How do I change the email subject and body?  
  **A:** You can do this by changing the contents of the subject.txt and body.txt files respectively
- **Q:** I don't like the font, how do I change it?
  **A:** simply replace the `.opensans.ttf` file. make sure to rename the file to `.opensans.ttf` or replace the variable name in the code.

- **Q:** The names of the participants are not properly aligned. What can I do?  
  **A:** Cert-Gen is manually drawing where the names are. To change the location of the names, find the x and y axis of where you want the name to be (you can do this in Photoshop or any other image editing software) and change the x and y axis in following code:

  ```
  draw.text(xy=(725, 680), text='{}'.format(
        row[1].value), fill=(255, 255, 255), font=font)
  ```

#### This software was created by James Jilhaney
