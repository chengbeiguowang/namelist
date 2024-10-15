import os


def get_project_application_name(work_dir) -> str:
    for file in os.listdir(work_dir):
        tail = os.path.splitext(file)[-1][1:]
        if file.__contains__("申报书") and tail == "pdf":
            return file

    print("Error in find program file: " + str(work_dir))
    return ''

