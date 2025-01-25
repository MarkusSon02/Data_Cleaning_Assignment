# Data Cleaning Project

## Overview

This project involves cleaning and preparing a dataset for analysis. The dataset contains information about students' participation in the presentation of other groups and their evaluations. The goal was to handle missing values, standardize columns, remove invalid data in order to answer three questions:

- Which students lied about participation?
- What are the participation marks of each student?
- What are the presentation marks of each group?

There are three Section in total (A03, A04, A06), and Section A03 and A06 have 10 Topics (1-10) and Section A04 have 9 Topics (1-10 excluding 3). Each group is a a Topic inside a Section, so there are in total 29 groups. Each student is identified by an Email Address. It is assumed that students can only attend group presentations from their own Sections. The Evaluate 1-5 columns are the marks that each student give for the presenting group, with the Score being the total, or presentation marks of the group given by the student.

---

## Steps to Clean the Data

### 1. **Loading the Data**

- The dataset is loaded from an Excel file using Pandas, turning it into a DataFrame object for easy processing:
  ```python
  data = pd.read_excel(file_path)
  ```
- All necessary columns are renamed for consistency and easier access:
  ```python
  data.columns = ["Timestamp", "Email Address", "Section", "Door Entry Number", "Topic", "Evaluate 1", "Evaluate 2", "Evaluate 3", "Evaluate 4", "Evaluate 5", "Comments", "Score"]
  ```
- `Timestamp` column is transformed into type datetime to ensure consistency:
  ```python
  data["Timestamp"] = pd.to_datetime(data["Timestamp"])
  ```

### 2. **Dropping Unnecessary Columns**

- Columns irrelevant to the analysis, such as `Comments`, are removed using:
  ```python
  data = data.drop("Comments", axis=1)
  ```

### 3. **Dropping invalid rows**

- Drop any rows with Topic 3 in Section A04:
  ```python
  data = data[~((data["Section"] == "Section A04, 2-3 pm on Monday and Wednesday") & (data["Topic"] == "Topic 3: Paying Employees to Relocate"))]
  ```


### 4. **Handling Missing Values**

- Rows with missing values in `Door Entry Number` column are manually checked by looking at the `Door Entry Number` column of other rows with the same `Email Address` value and created on the same day, and filled in:
  ```python
  data.loc[(data["Email Address"] == "10103@ualberta.ca") & (data["Topic"] == "Topic 3: Paying Employees to Relocate"), "Door Entry Number"] = 29
  ```
- Rows with missing values in `Section` column are manually checked by looking at the `Section` column of other rows with the same `Email Address` value, and filled in:
  ```python
  data.loc[(data["Email Address"] == "10109@ualberta.ca") & data["Section"].isna(), "Section"] = "Section A06, 3-4 pm on Monday and Wednesday"
  ```

### 5. **Handling Invalid Values**

- Certain rows have the wrong `Topic` value.
- This can be checked manually by looking at rows by the same `Email Address` on the same day.
- The `Section` value is then rectified:
  ```python
  data.loc[(data["Email Address"] == "1008@ualberta.ca") & (data["Door Entry Number"] == 43) & (data["Topic"] == "Topic 4: Labor Productivity During Recessions"), "Topic"] = "Topic 8: Brand-Name and Generic Drugs"
  ```
- A Boolean mask is used to identify invalid rows and extract them for further inspection.

### 6. **Identifying students who lied about participation and removing them from the original DataFrame**

- The type of `Door Entry Number` column is forcefully changed to float, turning any string values into NaN to be removed:
  ```python
  data["Door Entry Number"] = pd.to_numeric(data["Door Entry Number"], errors="coerce")
  ```
- The data DataFrame is then sorted by `Timestamp` in ascending order. This is to identify any student that entered a `Door Entry Number` that has already been entered.
- These students will automatically recorded and removed from the `data` DataFrame:
  ```python
  data = data.sort_values(by="Timestamp")
  lied_about_participation_df_1 = data[data.duplicated(subset=["Section", "Door Entry Number", "Topic"], keep="first")]
  data = data.drop_duplicates(subset=["Section", "Door Entry Number", "Topic"], keep="first")
  ```
- A Boolean mask is then used to identify invalid rows and extract them as students who also lied about participation.
- This mask is then used to seperate the invalid rows from the `data` DataFrame:
  ```python
  invalid_mask = (data["Door Entry Number"].isna()) | (data["Door Entry Number"] < 1) | (data["Door Entry Number"] > 60)
  lied_about_participation_df_2 = data[invalid_mask]
  data = data[~invalid_mask]
  ```
- The two DataFrame recording the students who lied about participation are then combined into one DataFrame, later to be saved as an Excel file:
  ```python
  invalid_mask = (data["Door Entry Number"].isna()) | (data["Door Entry Number"] < 1) | (data["Door Entry Number"] > 60)
  lied_about_participation_df_2 = data[invalid_mask]
  data = data[~invalid_mask]
  ```

### 7. **Calculating Participation Marks**

- Rows that are valid are then aggregated on the `Email Address` column and then counted to determine the participation marks for each student.
  ```python
  email_counts = data.groupby(["Email Address", "Section"]).size().reset_index(name="Count")
  email_counts["Participation Mark"] = email_counts["Count"] * 0.5
  email_counts.loc[email_counts["Section"] == "Section A04, 2-3 pm on Monday and Wednesday", "Participation Mark"] += 0.5
  participation_marks_df = email_counts.drop("Count", axis=1)
  ```
- Each participation will contribute 0.5 to the overall participation marks. 
- Students in Section A04 only had 9 Topics compared to 10 from the other Sections. Therefore, in order to ensure fairness, they are rewared 0.5 marks.


### 8. **Standardizing "Evaluate 1-5 "Columns Values**

- The `Evaluate` columns often contained fractional values like `8/10`. This function was written in order to be applied to each cell for standardization:
  ```python
  def convert_to_score(value):
      try:
          numerator, denominator = map(float, value.split("/"))
          return (numerator / denominator) * 2
      except:
          return value
  ```
- In order clean the data, each value is split by the space, and the part before the space is used.
- Then the `convert_to_score` function is applied to each cell. Cells that are not fractional values will not be affected.
- Then the values are forcefully converted to type float.
- And any value that are invalid, such as strings that do not have numbers at the beginning or removed. This is because the number of values that are strings only make up 1-2% of the dataset, and will not affect the data representation of the population.
  ```python
  evaluate_columns = ["Evaluate 1", "Evaluate 2", "Evaluate 3", "Evaluate 4", "Evaluate 5"]
  for col in evaluate_columns:
    data[col] = data[col].fillna("").astype(str).str.split().str[0]
    data[col] = data[col].apply(calculate_score)
    data[col] = pd.to_numeric(data[col], errors="coerce")
    data.dropna(subset=[col])
  ```

### 9. **Calculating Presentation Marks for Each Group**

- `Evaluate` columns or summed for each student:
  ```python
  data["Score"] = data[evaluate_columns].sum(axis=1)
  ```
- The data is then aggregated based on `Section` and `Topic`, representing groups, and applied the mean aggregate function to calculate the final marks for each group:
  ```python
  columns_to_mean = evaluate_columns + ["Score"]
  presentation_marks_df = data.groupby(["Section", "Topic"])[columns_to_mean].mean().reset_index()
  ```

---

## Outputs

- **Data Table of Students who Lied about Participation**: The DataFrame is saved as a Excel file in the folder `results`:
  ```python
  lied_about_participation_df.to_excel("results/Students_Who_Lied_About_Participation.xlsx")
  ```
- **Data Table of Participation Marks**: The DataFrame is saved as a Excel file in the folder `results`:
  ```python
  participation_marks_df.to_excel("results/Participation_Marks.xlsx")
  ```
- **Data Table of Presentation Marks**: The DataFrame is saved as a Excel file in the folder `results`:
  ```python
  presentation_marks_df.to_excel("results/Presentation_Marks.xlsx")
  ```

---

## Key Challenges

- **Dealing with Mixed Data Types**: Columns contained inconsistent formats (e.g., fractions, numeric values, text). These were handled using custom functions to standardize values.
- **Handling Missing**: Missing values and outliers were manually checked and replaced or removed to ensure data integrity.
- **Handling Invalid Data**: Falsified data such as certain values for `Door Entry Number` were difficulty to identified and required careful manual inspection. Other falsified values of `Door Entry Number` were able to be identified automatically using code.

---

## Dependencies

- Python 3.7+
- Pandas
- OpenPyXL (for Excel support)

---

## Author

Son Tran

