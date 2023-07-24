import docx.text.font

from github import Github
import docx
from docx.shared import Pt, Inches
from github import Auth
import requests
import pypandoc
import pandoc
import re
from pathlib import Path

# To use this program, go into main and change the repo_name to the desired repository and its creator


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


def convert_gfm_to_docx(file_name: Path, output_file: Path):
    # with open("markdown_test.md", "w") as file:
    #     file.writelines(pypandoc.convert_text(text, "markdown", format="gfm"))
    pypandoc.convert_file(
        file_name.as_posix(), to="docx", outputfile=output_file.as_posix()
    )


headers = {
    "Authorization": f"Bearer {AUTH_TOKEN}",
    "X-GitHub-Api-Version": "2022-11-28",
}


# gets comments regarding issue, helper function for get_issues
def get_comments(issue):
    comments_url = issue.comments_url
    response = requests.get(comments_url, headers=headers)
    json_data = response.json()

    comments = []
    for i in range(len(json_data)):
        comment_tuple = (
            str(json_data[i].get("user")["login"]),
            str(json_data[i].get("body")),
        )
        comments.append(comment_tuple)
    return comments


# gets issues from GitHub
def get_issues(repository):
    # gets repository
    repo = g.get_repo(repository)

    # stores all issues
    open_issues = repo.get_issues(state="open")

    issues_list = []
    for issue in open_issues:
        if not issue.pull_request:
            if issue.labels:
                issue_dict = {
                    "id_num": issue.number,
                    "title": issue.title,
                    "body": issue.body,
                    "user": issue.user.login,
                    "assignees": issue.assignees,
                    "published": issue.created_at.isoformat(),
                    "labels": issue.labels[0].name,
                    "milestone": issue.milestone,
                    "comments": None,
                }
            else:
                issue_dict = {
                    "id_num": issue.number,
                    "title": issue.title.replace("`", "'"),
                    "body": issue.body,
                    "user": issue.user.login,
                    "assignees": issue.assignees,
                    "published": issue.created_at.isoformat(),
                    "labels": None,
                    "milestone": issue.milestone,
                    "comments": None,
                }

            if len(issue.assignees) < 1:
                issue_dict["assignees"] = None

            if issue.comments != 0:
                issue_dict["comments"] = get_comments(issue)

            issues_list.append(issue_dict)

    return issues_list


def create_md_doc(issue_list, output_dir):
    md_dir = output_dir / "markdown"
    md_dir.mkdir(parents=True, exist_ok=True)

    for issue in issue_list:
        output_name = md_dir / f"issue_{issue['id_num']}.md"

        with open(output_name, "w") as file:
            file.write(f"<center>**{issue['title']}**</center>\n")

            file.write(
                "*Issue No. "
                + str(issue["id_num"])
                + " opened by "
                + str(issue["user"])
                + " on "
                + str(issue["published"].replace("T", " at "))
                + "    Type: "
                + str(issue["labels"])[5:]
                + "*\n\n"
            )
            file.write("\n### Issue:\n")
            file.writelines(
                pypandoc.convert_text(issue["body"], "markdown", format="gfm")
            )
            file.write("\n\n### Additional Information: \n")
            file.write("* Milestones: " + str(issue["milestone"]) + "\n")
            file.write("* Assignees: " + str(issue["assignees"]) + "\n")

            file.write("\n\n### Comments: \n")
            if issue["comments"]:
                for comment in issue["comments"]:
                    file.write("\n#### " + comment[0] + ":\n")
                    file.writelines(
                        pypandoc.convert_text(comment[1], "markdown", format="gfm")
                    )
            else:
                file.write("\nNo comments at the moment!\n")

            file.write("<br/>\n")
            file.close()


def convert_md_folder(out_dir: Path):
    word_dir = out_dir / "docx"
    md_dir = out_dir / "markdown"
    word_dir.mkdir(parents=True, exist_ok=True)
    for file in md_dir.glob("*.md"):
        assert file.is_file()
        out_file_name = word_dir / f"{file.stem}.docx"
        convert_gfm_to_docx(file, out_file_name)
        format_word_doc(out_file_name)


def format_word_doc(file_name):
    document = docx.Document(file_name)

    heading_para = document.add_heading("\tIssues in The " + " Repository", 0)

    # header_style = document.styles["Heading 1"]
    # header_font = header_style.font
    # header_font.size = Pt(18)
    #
    # section = document.sections[0]
    # heading = section.header
    # heading_para = heading.paragraphs[0]
    # heading_para.style = header_style
    # heading_para.text = "\tIssues in The " + " Repository"

    logo = heading_para.add_run()
    logo.add_picture("gen_company_logo.png", width=Inches(1))


if __name__ == "__main__":
    pypandoc.download_pandoc()
    repo_name = "Project-MONAI/MONAILabel"

    issues = get_issues(repo_name)

    out_dir = Path(repo_name + "_issues")
    out_dir.mkdir(parents=True, exist_ok=True)

    create_md_doc(issues[:10], out_dir)
    convert_md_folder(out_dir)
    # pypandoc.convert_file("markdown_test.md", to="docx", outputfile="test.docx")

    print("finished")
