# excel_utils

## Getting Started

Set up and run the project with these steps:

1. **Install Python 3**  
    Download and install Python 3 from [python.org](https://www.python.org/).

2. **Install Visual Studio Code**  
    Get [Visual Studio Code](https://code.visualstudio.com/) for a powerful Python development experience.

3. **Set Up a Virtual Environment**  
    Ensure Python 3.13+ is installed:
    ```bash
    python3 --version
    ```
    Create and activate a virtual environment:
    ```bash
    python3 -m venv .venv
    source .venv/bin/activate  # On Windows: .venv\Scripts\activate
    ```

4. **Install Dependencies**  
    Install required packages:
    ```bash
    pip install -r requirements.txt
    ```

5. **Prepare Comparison Folders**  
    Create the following structure:
    ```
    files/number/1
    files/number/2
    ```
    Place files to compare in `files/number/1` and `files/number/2`.

6. **Configure the Project**  
    Copy `config.py.example` to `config.py` and update settings as needed.

7. **Run the Application**  
    Start the main script:
    ```bash
    python3 main.py
    ```
