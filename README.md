

# Campus EduTrack 📁  🐍  📊

Campus EduTrack is a Python-based application designed to automate student data management, streamlining processes, and improving accuracy at SIES Graduate School of Technology, Nerul, Navi Mumbai  🏫  💻. This project leverages Python 🐍 and the `xlwings` library to handle student data and generate processed Mentee Information Sheet efficiently  📑.

## Features

* Automates student data processing, including personal details and academic records  🧑‍🎓  📋.
* Generates output files for individual students in a structured and organized manner  📁  🗃️.
* Eco-friendly by eliminating the need for paper-based record-keeping  🌳.
* Intuitive GUI for selecting database files, defining row ranges, and executing operations  🖱️  🧭  ⚙️.

## Prerequisites

* Python 3.6 or later  🐍
* `xlwings` library (Install using `pip install xlwings`)  📦
* Microsoft Excel (Office 2021 or later recommended)  📊

## Installation

1. Create a folder for the project  📁.
2. Download the following files from the repository:
   * `Database.xlsx` 📑
   * `Campus_Edu.py` 🐍
   * `template.xlsx`  📑
3. Place these files into the created folder  📁.
4. Alternatively, clone the repository using:

```bash
git clone https://github.com/png-bash/CampusEduTrack.git
```

## How to Run

1. Navigate to the project folder in your terminal  🧭
2. Run the program with:

```bash
python Campus_Edu.py
```

3. The program will create an `output` folder in the current directory for saving processed student files  📁  💾.

## Using the Application

1. **GUI**: After running the script the graphical user interface is launched  🖱️.
2. **Select Database**:
   * Click the "Browse Database" button  🖱️  📁.
   * Select the `Database.xlsx` file from your directory  📁.
3. **Set Row Range**:
   * Input the starting row (e.g., `3`) and ending row (e.g., `4`) to process specific students  🔢.
   * Row `3` corresponds to the first row of student data in the sample database  📋.
4. **Select Operation**:
   * Use the "Select Operation" dropdown to choose a task  🖱️  dropdown.
   * Default: "Fill Page 1 and Page 2" (When running the software for the first time for a student, it is recommended to keep the default option).
5. **Process Files**:
   * Click "Process Selected Operation."  🖱️
   * The status updates in the GUI, and a final alert confirms when processing is complete  ✅.
6. **View Output**:
   * Processed files are saved in the `output` folder, named by student roll number  📁.

## Generating Additional Pages

To process other pages, such as Semester 1 data:

1. Change the operation in the "Select Operation" dropdown  🖱️  dropdown.
2. Repeat the steps for selecting the database and defining the row range  🔁.
3. Click "Process Selected Operation" to generate the required files  🖱️  💾.

## Sample Data

The provided `Database.xlsx` includes sample data for 10 students to test the program's functionality  📋.

## Notes

* Follow the sequence of operations in the dropdown for accurate file generation.
* Ensure all required files (`Database.xlsx`, `Campus_Edu.py`, `template.xlsx`) are in the same folder before running the script  📁  ⚠️.
* Output files are organized in the `output` folder by student roll number  📁.

## Contributing

Contributions are welcome! If you encounter issues or have ideas for improvement, feel free to submit an issue or a pull request to the repository  🤝  🔧.

## Acknowledgments

* Developed using Python and the `xlwings` library  🐍  📊.
* Inspired by the need for efficient, eco-friendly student data management  💡  🌳.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details  📄.

## Contact

For questions, suggestions, or feedback, please contact Prathamesh Gajare (gajareprathamesh@gmail.com)  📧.
