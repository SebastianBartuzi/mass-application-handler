import subprocess


def install_packages():
    try:
        # Run the pip install command
        subprocess.check_call(['pip', 'install', 'docx2pdf', 'python-docx', 'openpyxl', 'pywin32'])
        print("Packages installed successfully!")
    except subprocess.CalledProcessError as e:
        print("An error occurred while installing packages:", e)


if __name__ == "__main__":
    install_packages()
