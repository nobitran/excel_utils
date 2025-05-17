# excel_utils

## Getting Started

Follow these steps to set up and run the project:

1. **Install Python 3**  
    Ensure you have Python 3 installed on your system. You can download it from [python.org](https://www.python.org/).

2. **Install Visual Studio Code (VS Code)**  
    Download and install [Visual Studio Code](https://code.visualstudio.com/). It is a lightweight and powerful code editor that works well with Python projects.

3. **Set up a virtual environment**  
    Check current Python3 version (required 3.13 ex: Python 3.13.2)
    ```
    python3 --version
    ```

    Create and activate a virtual environment in the project directory:
    ```bash
    python3 -m venv .venv
    source .venv/bin/activate  
    # On Windows, use .venv\Scripts\activate
    ```

4. **Install dependencies**  
    Install the required packages using `requirements.txt`:
    ```bash
    pip3 install -r requirements.txt
    ```
5. **Set up comparison folders**  
    Create the following folder structure to organize files for comparison:
    ```
    files/number/1
    files/number/2
    ```
    Place the files you want to compare inside `files/number/1` and `files/number/2` respectively.

6. **Create a configuration file**  
    Create a file named `config.py` from `config.py.example` in the project directory . Add any necessary configuration settings to this file.

7. **Run the application**  
    Execute the main script to start the application:
    ```bash
    python3 main.py
    ```