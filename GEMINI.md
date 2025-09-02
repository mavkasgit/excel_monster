# Project Overview

This project is a VBA application for Microsoft Excel designed to manage a task list. It uses a UserForm to provide a graphical interface for viewing and interacting with tasks defined in an Excel worksheet.

The project is structured into three VBA modules:
*   **Module1**: Contains the entry point to launch the application's UserForm.
*   **Module2**: Implements the core logic for processing and executing tasks.
*   **ModuleForm**: Defines the UserForm interface and its event handlers.

The `Rules.md` file contains a set of coding standards and best practices for the VBA code in this project.

# Building and Running

To run this application:
1.  Open the Excel file containing the VBA project.
2.  Open the VBA Editor (Alt+F11).
3.  Run the `ShowForm` subroutine in `Module1`.

This will display the main UserForm, which loads tasks from the "План" worksheet.

# Development Conventions

The `Rules.md` file outlines the development conventions for this project. Key conventions include:
*   All code must be written in VBA and run in the VBA Editor without external libraries.
*   Comments are mandatory to explain the logic of the code.
*   Code should be well-structured with proper indentation.
*   `Option Explicit` must be used to prevent errors from implicit variable declarations.
*   Code should be modular, with each procedure performing a single task.
*   Avoid using `.Select` and `.Activate`.
*   Use error handling (`On Error GoTo`) to prevent crashes.
