import docx.text.font
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

from github import Github
import docx
from docx.shared import Pt, Inches
from github import Auth
import requests
import pypandoc
import pandoc
import re
import os
from Markdown2docx import Markdown2docx
# project = Markdown2docx('README')
# project.eat_soup()

# TODO Provide instructions for creating Auth token
# log into github and go to this page https://github.com/settings/tokens
# get token and type into auth_token.txt file
# TODO Read in Auth token from txt file
f = open("auth_token.txt")
AUTH_TOKEN = f.readline()[:-1]

# using an access token
auth = Auth.Token(AUTH_TOKEN)

# First create a GitHub instance:

# Public Web GitHub
g = Github(auth=auth)

# strips GitHub issue template, gets rid of newline and return characters, and gets image links

def text_trimmer(text):
    text_dict = {"text": text, "code": [], "image_urls": []}
    # text.replace(">", "\t")
    text.replace('\\', "")

    while "```" in text:
        code_index = text.index("```")
        end_code_index = text.find("```", code_index + 3) + 3
        if end_code_index > code_index:  # catches case where there is no ending "```" and any other oddity with the end_code_index
            code_snippet = text[code_index: end_code_index]
            text_dict["code"].append(code_snippet)
            text = re.sub(r'```[\s\S]*?```', "\n [CODE REPLACE] \n", text, 1)
            # print(f"found {len(text_dict['code'])} code")
        else:
            break

    while "https://github.com/Project-MONAI/MONAILabel/assets/" in text:
        # print("1 image")
        start_link_index = text.index("https://github.com/Project-MONAI/MONAILabel/assets/")
        end_link_index = text.find(")", start_link_index)
        image_url = text[start_link_index: end_link_index]
        text_dict["image_urls"].append(image_url)
        if "![" in text and "](" in text:
            text = re.sub(r'!\[.*?\]\((.*?)\)', "\n [IMAGE REPLACE] \n", text, 1)  # check for html tags
        elif "<img" in text:
            text = re.sub(r'<img[^>]+>', "\n [IMAGE REPLACE] \n", text, 1)  # check for html tags
        # if len(text_dict['image_urls']) > 2:
        #     print("here")
        #
        # print(f"found {len(text_dict['image_urls'])} images")
    text_dict["text"] = text

    return text_dict

def ordering_list(comment_dict):
    text = comment_dict["text"]
    code_rp_str = "[CODE REPLACE]"
    image_rp_str = "[IMAGE REPLACE]"
    ordered_rp_str = "[ORDERED REPLACE]"
    ordered_list = []
    image_num = 0
    code_num = 0
    current_index = 0
    while code_rp_str in text or image_rp_str in text:
        code_index = text.find(code_rp_str, current_index)
        image_index = text.find(image_rp_str, current_index)

        if (code_index < image_index and code_index > 0) or (image_index < 0 and code_index > 0):
            ordered_list.append(comment_dict["code"][code_num])
            code_num += 1
            current_index = code_index + len(ordered_rp_str)
            text = text.replace(code_rp_str, ordered_rp_str, 1)
        elif (code_index > image_index and image_index > 0) or (image_index > 0 and code_index < 0):
            ordered_list.append(comment_dict["image_urls"][image_num])
            image_num += 1
            current_index = image_index + len(ordered_rp_str)
            text = text.replace(image_rp_str, ordered_rp_str, 1)
        else:
            print("Something went wrong")

    comment_dict["text"] = text
    comment_dict["ordered_list"] = ordered_list

    return comment_dict

def convert_gh_to_rtf(text):
    with open("markdown_test.md", "w") as file:
        file.writelines(pypandoc.convert_text(text, "markdown", format="gfm"))

    return pypandoc.convert_text(text, "markdown", format="gfm")  # instead of using markdown_github was given warning to use gfm(github flavored markdown)

headers = {"Authorization": f"Bearer {AUTH_TOKEN}", "X-GitHub-Api-Version": "2022-11-28"}

# gets comments regarding issue, helper function for get_issues
def get_comments(issue):
    comments_url = issue.comments_url
    response = requests.get(comments_url, headers=headers)
    json_data = response.json()

    comments = []
    for i in range(len(json_data)):
        # comment = text_trimmer(str(json_data[i].get("body")))
        comment_tuple = (str(json_data[i].get("user")["login"]), str(json_data[i].get("body")))
        comments.append(comment_tuple)
    return comments

# gets issues from GitHub
def get_issues(repository):
    # gets repository
    repo = g.get_repo(repository)

    # stores all issues
    open_issues = repo.get_issues(state="open")

    issues_list = []
    i = 0
    for issue in open_issues:
        if not issue.pull_request:
            if issue.labels:
                issue_dict = {"id_num": issue.number, "title": issue.title, "body": issue.body, "user": issue.user.login, "assignees": issue.assignees,
                              "published": issue.created_at.isoformat(), "labels": issue.labels[0].name, "milestone": issue.milestone, "comments": None}
            else:
                issue_dict = {"id_num": issue.number, "title": issue.title.replace("`", "'"), "body": issue.body, "user": issue.user.login, "assignees": issue.assignees,
                              "published": issue.created_at.isoformat(), "labels": None, "milestone": issue.milestone, "comments": None}

            if issue.comments != 0:
                issue_dict["comments"] = get_comments(issue)

            if issue_dict["body"]:
                # body_dict = text_trimmer(issue_dict["body"])
                # body_dict = ordering_list(body_dict)
                body_dict = issue_dict["body"]
                body_dict = convert_gh_to_rtf(body_dict)
                issue_dict["body"] = body_dict

            i += 1
            issues_list.append(issue_dict)

    return issues_list

def create_md_doc(issue_list, repo_name):
    with open("markdown_test.md", "w") as file:
        for issue in issue_list:
            file.writelines(pypandoc.convert_text(issue["body"], "markdown", format="gfm"))

            file.write("\n# Comments : \n")
            if issue["comments"]:
                for comment in issue["comments"]:
                    file.write("\n## " + comment[0] +" :  \n")
                    file.writelines(pypandoc.convert_text(comment[1], "markdown", format="gfm"))
            else:
                file.write("\n No comments at the moment!  \n")

            file.write("<br/>")
            file.write("<br/>")
            file.write("<br/>")


# creates word document
def create_word_doc(issue_list, repo_name):
    # create document
    document = docx.Document()

    # creates styles for header, heading and the normal paragraph
    header_style = document.styles['Header']
    header_font = header_style.font
    header_font.size = Pt(18)

    heading_style = document.styles['Heading 3']
    heading_font = heading_style.font
    heading_font.size = Pt(14)

    style = document.styles['Normal']
    font = style.font
    font.size = Pt(10)

    # make cover page
    main_header = document.add_heading(repo_name, 0)
    main_header.alignment = 1
    main_header.paragraph_format.space_after = Pt(28)  # adds spaces after end of paragraph
    main_header.paragraph_format.space_before = Pt(16)

    front_para = document.add_paragraph(str(len(issue_list)) + " issues gathered from the github repository regarding " + repo_name[repo_name.index("/") + 1:])
    front_para.style = document.styles['Header']
    front_para.alignment = 1

    document.add_page_break()

    # creates page heading
    section = document.sections[0]
    heading = section.header
    heading_para = heading.paragraphs[0]
    heading_para.style = document.styles['Heading 3']
    heading_para.text = "\tIssues in The " + repo_name[repo_name.index("/") + 1:] + " Repository"

    logo = heading_para.add_run()
    logo.add_picture("gen_company_logo.png", width=Inches(1))

    for issue in issue_list:
        header = document.add_heading(issue["title"], 0)  # creates header with title of issue as content
        header.paragraph_format.space_after = Pt(10)
        header.paragraph_format.space_before = Pt(16)
        header.style = document.styles["Header"]
        header.alignment = 1

        # the info paragraph contains the issue number, its author, publish date, and type of issue
        info_paragraph = document.add_paragraph(style="Normal")
        if issue["labels"]:
            info_paragraph.add_run("Issue No. " + str(issue["id_num"]) + " opened by " + str(issue["user"]) + " on " + str(issue["published"].replace("T", " at "))
                                   + "    Type: " + str(issue["labels"])[5:]).bold = True
        else:
            info_paragraph.add_run("Issue No. " + str(issue["id_num"]) + " opened by " + str(issue["user"]) + " on " + str(issue["published"].replace("T", " at "))
                                   + "  Type: None").bold = True
        info_paragraph.paragraph_format.space_after = Pt(14)

        # the body paragraph contains the body of the issue message
        body_paragraph = document.add_paragraph(style="Body Text")
        body_text = issue["body"]
        body_paragraph.add_run(body_text)
        # replace_flag = "[ORDERED REPLACE]"
        # if len(issue["body"]["ordered_list"]):
        #     print("Here")
        #     for item in issue["body"]["ordered_list"]:
        #         replace_index = body_text.index(replace_flag)
        #         if item[:2] == "```":
        #             body_paragraph.add_run(body_text[:replace_index] + "\n")
        #             body_text = body_text[replace_index:]
        #             table = document.add_table(rows=1, cols=1)
        #             cell = table.cell(0, 0)
        #             cell.text = item[3:len(item) - 3]
                    # insertion

                #     issue["body"]["text"] = issue["body"]["text"]
                # else:
                #     body_paragraph.add_run(issue["body"]["text"])
                #     # insertion
                #     issue["body"]["text"] = issue["body"]["text"]
        #
        # else:
        #     body_paragraph = document.add_paragraph(issue["body"]["text"], style="Body Text")
        body_paragraph.style = document.styles['Normal']
        body_paragraph.paragraph_format.space_after = Pt(16)

        # the additional info has info on milestones and the assignees
        additional_info = document.add_paragraph(style="Normal")
        additional_info.add_run("Additional Information: ").bold = True

        # conditionals are used because of NoneType errors
        if issue["milestone"]:
            document.add_paragraph("Milestones: " + str(issue["milestone"].title), style="List Bullet")
        else:
            document.add_paragraph("Milestones: None", style="List Bullet")

        if issue["assignees"]:
            document.add_paragraph("Assignees: " + ", ".join([x.login for x in issue["assignees"]]), style="List Bullet")
        else:
            document.add_paragraph("Assignees: None", style="List Bullet")

        comment_paragraph = document.add_paragraph(style="Normal")
        if issue["comments"]:
            comment_paragraph = document.add_paragraph(style="Normal")
            for comment in issue["comments"]:
                comment_paragraph.add_run(comment[0] + ": \n").bold = True
                comment_paragraph.add_run(comment[1] + "\n")
                comment_paragraph.add_run("-" * 120 + "\n")
                comment_paragraph.paragraph_format.space_after = Pt(14)
        else:
            comment_paragraph.add_run("No comments at the moment!")

        document.add_page_break()  # ends the page at this point for the next issue

    document.save("demo_issues.docx")


if __name__ == '__main__':
    pypandoc.download_pandoc()
    repo_name = "InsightSoftwareConsortium/ITK"

    issues = get_issues(repo_name)
    create_md_doc(issues[:10], repo_name)
    pypandoc.convert_file("markdown_test.md", to='docx', outputfile="test.docx")

    # create_word_doc(issues[:20], repo_name)

    print("finished")

# Check to see if can directly convert body markdown to docx
