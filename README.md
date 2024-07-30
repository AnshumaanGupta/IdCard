# ID Card Generation Script

This Python script generates ID cards by reading data from an Excel file. It uses the Pillow and python-barcode libraries to create front and back images of ID cards and saves them as PDF files.

## Features

- Reads data from an Excel file
- Generates barcode images
- Creates front and back images of ID cards
- Saves the generated ID cards as PDF files
- Logs skipped rows due to errors

## Prerequisites

- Python 3.x
- Pillow 9.5.0
- python-barcode
- pandas

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/id-card-generator.git
   ```
2. Navigate to the project directory:
   ```bash
   cd id-card-generator
   ```
3. Install the required Python packages:
   ```bash
   pip install pillow==9.5.0
   pip install python-barcode
   pip install pandas
   ```

## Usage

1. Prepare your Excel file (`data.xlsx`) with the following columns:

   - `name`
   - `old_rollno`
   - `new_rollno`
   - `branch`
   - `mobileno`
   - `email`
   - `parentname`
   - `address`
   - `codeid`
   - `imagepath`
   - `validity`

2. Ensure you have your ID card templates (`card1.png` and `card2.png`) and student photos in the `studentphoto/` directory.

3. Run the script:

   ```bash
   python id_card_generator.py
   ```

4. The generated ID cards will be saved in the `output/` directory as images and in the `pdf/` directory as PDF files.

## Logging

- The script logs any skipped rows due to errors in `skipped_rows.log`.

## Example

```python
# Sample data row in the Excel file
name: John Doe
old_rollno: 123456
new_rollno: 654321
branch: Computer Science
mobileno: 1234567890
email: johndoe@example.com
parentname: Jane Doe
address: 123 Main St, Anytown, USA
codeid: 123456
imagepath: johndoe.png
validity: 2025-12-31
```
