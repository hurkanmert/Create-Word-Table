## Overview

Optimized frequency and geometry tables of selected peptides are created using the Gaussian program in Atomic and Molecular Physics. Eliminating the necessary data in these extremely long tables constitutes a time-consuming task. Therefore, using the Python programming language, a table can be created in a Word document if the desired range properties are specified.

## ✨ Features

- **Variable Multipliers:** The double scaling factor has a specific coefficient on it and justified values for different multipliers. You can change the values and limit of these multipliers.


- **Delimiting Parameters:** In order to capture the necessary data sets in the sample files included in the study, I selected the closest word on the data sets as the delimiter. You can change this variable to limit it on different files.

## Getting Started

### Installation

Clone the repository and install the required dependencies:

```bash
pip install -r requirements.txt

```

### Usage

Run the main script to create table:

```bash
python main.py
```

## ⚙️ Configuration

sep_arg_one and sep_arg_two variables are required to limit the frequency log file. Likewise, sep_arg_three and sep_arg_four are used as necessary variables to limit the geometry log file.

```bash
sep_arg_one = "normal coordinates:"
sep_arg_two = "Thermochemistry"
sep_arg_three = "Optimized Parameters"
sep_arg_four = "Stoichiometry"
```

## 🤘🏻 rock n roll

If you have questions or would like more information, please feel free to email me.<br>
<b>hurkanmertd@gmail.com</b>