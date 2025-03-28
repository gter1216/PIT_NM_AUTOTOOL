# AutoTool

UDS AutoTool is a tool designed for the PIT team to automate UDS test works.

## Features
- Feature 1: Description of feature 1.
- Feature 2: Description of feature 2.
- Feature 3: Description of feature 3.

## Installation
To install the dependencies, run:
```bash
pip install -r requirements.txt
```

## Usage
To start the tool, run:
```bash
python main.py
```

### Script Usage Details
- **CURR_REL**: This directory should contain the latest diagnostic configuration tables released by the company.
- **LAST_REL**: This directory contains the test parameter tables generated based on the previous version of the diagnostic configuration tables.
- **OUTPUT**: This directory will store the test parameter tables generated based on the latest diagnostic configuration tables.

The script will compare the new test parameter tables with the old ones and print out which ECUs have been added and which have been removed.

Ensure that the directories are correctly set up before running the script to get accurate results.

### How to Use the Script
1. **Set Up Directories**:
   - Place the latest diagnostic configuration tables in the `CURR_REL` directory.
   - Ensure the `LAST_REL` directory contains the previous version's test parameter tables.
   - The `OUTPUT` directory will be used to store the newly generated test parameter tables.

2. **Run the Script**:
   - Open a command line interface.
   - Execute the following command to start the tool:
     ```bash
     python main.py
     ```

3. **Script Functionality**:
   - The script will analyze the differences between the new and old test parameter tables.
   - It will output a list of ECUs that have been added or removed in the new version.

Make sure all directories and files are correctly prepared before running the script to ensure accurate results.

### Script Parameters
The script accepts the following parameter:
- `--vehicle`: Specifies the vehicle type (e.g., BLANC_RL201, CETUS_RL201).

Example command:
```bash
python main.py --vehicle BLANC_RL201
```

Ensure that the vehicle type provided is correct and supported by the script.

## Contributing
1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Commit your changes (`git commit -am 'Add new feature'`).
4. Push to the branch (`git push origin feature-branch`).
5. Create a new Pull Request.

## License
This project is licensed under the MIT License.

## Author
Author: Charles Xu
