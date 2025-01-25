import pandas as pd

# Load the dataset
file_path = "Case Presentation Evaluation data - BUEC 420.xlsx"  # Replace with the actual file path
data = pd.read_excel(file_path)

# Rename columns and dropping comments column
data.columns = ["Timestamp", "Email Address", "Section", "Door Entry Number", "Topic", "Evaluate 1", "Evaluate 2", "Evaluate 3", "Evaluate 4", "Evaluate 5", "Comments", "Score"]
data = data.drop("Comments", axis=1)

# Ensure Timestamp column is in datetime format
data["Timestamp"] = pd.to_datetime(data["Timestamp"])

# Drop exact duplicates
data = data.drop_duplicates()

# Drop any rows with Topic 3 in Section A04 
data = data[~((data["Section"] == "Section A04, 2-3 pm on Monday and Wednesday") & (data["Topic"] == "Topic 3: Paying Employees to Relocate"))]

# Filling in missing values for Door Entry Number column
data.loc[(data["Email Address"] == "10103@ualberta.ca") & (data["Topic"] == "Topic 3: Paying Employees to Relocate"), "Door Entry Number"] = 29
data.loc[(data["Email Address"] == "10134@ualberta.ca") & (data["Topic"] == "Topic 8: Brand-Name and Generic Drugs"), "Door Entry Number"] = 22
data.loc[(data["Email Address"] == "10138@ualberta.ca") & (data["Door Entry Number"] == "41, I had submitted a form for topic 4 without door entry number by mistake."), "Door Entry Number"] = 41
data.loc[(data["Email Address"] == "10138@ualberta.ca") & (data["Door Entry Number"].isna()), "Door Entry Number"] = 41
data.loc[(data["Email Address"] == "10112@ualberta.ca") & (data["Door Entry Number"] == "15i"), "Door Entry Number"] = 15
data.loc[(data["Email Address"] == "1016@ualberta.ca") & (data["Topic"] == "Topic 2: Estimating the Effect of an iTunes Price Change"), "Door Entry Number"] = 11
data.loc[(data["Email Address"] == "1016@ualberta.ca") & (data["Topic"] == "Topic 4: Labor Productivity During Recessions"), "Door Entry Number"] = 50
data.loc[(data["Email Address"] == "1021@ualberta.ca") & (data["Topic"] == "Topic 4: Labor Productivity During Recessions"), "Door Entry Number"] = 25
data.loc[(data["Email Address"] == "1028@ualberta.ca") & (data["Topic"] == "Topic 9: Sale Prices"), "Door Entry Number"] = 19
data.loc[(data["Email Address"] == "1069@ualberta.ca") & (data["Topic"] == "Topic 9: Sale Prices"), "Door Entry Number"] = 54
data.loc[(data["Email Address"] == "1088@ualberta.ca") & (data["Topic"] == "Topic 8: Brand-Name and Generic Drugs"), "Door Entry Number"] = 59

# Filling in missing values for Section column
data.loc[(data["Email Address"] == "10109@ualberta.ca") & data["Section"].isna(), "Section"] = "Section A06, 3-4 pm on Monday and Wednesday"
data.loc[(data["Email Address"] == "10147@ualberta.ca") & data["Section"].isna(), "Section"] = "Section A04, 2-3 pm on Monday and Wednesday"
data.loc[(data["Email Address"] == "1099@ualberta.ca") & data["Section"].isna(), "Section"] = "Section A06, 3-4 pm on Monday and Wednesday"

# Fixing certain rows where the Topic was entered incorrectly
data.loc[(data["Email Address"] == "1008@ualberta.ca") & (data["Door Entry Number"] == 43) & (data["Topic"] == "Topic 4: Labor Productivity During Recessions"), "Topic"] = "Topic 8: Brand-Name and Generic Drugs"
data.loc[(data["Email Address"] == "1056@ualberta.ca") & (data["Door Entry Number"] == 59) & (data["Topic"] == "Topic 8: Brand-Name and Generic Drugs"), "Topic"] = "Topic 9: Sale Prices"


# Drop rows where the door entry number is not between 1 and 60
data["Door Entry Number"] = pd.to_numeric(data["Door Entry Number"], errors="coerce")
data = data.sort_values(by="Timestamp") # Sort by Timestamp in ascending order to drop the rows with the same Door Entry Number as another student but at a later date
lied_about_participation_df_1 = data[data.duplicated(subset=["Section", "Door Entry Number", "Topic"], keep="first")] # Save those students to add to the students who lied list
data = data.drop_duplicates(subset=["Section", "Door Entry Number", "Topic"], keep="first")
invalid_mask = (data["Door Entry Number"].isna()) | (data["Door Entry Number"] < 1) | (data["Door Entry Number"] > 60)
lied_about_participation_df_2 = data[invalid_mask] # Apply the mask to retrieve a table of students who lied about door number entry
lied_about_participation_df = pd.concat([lied_about_participation_df_1, lied_about_participation_df_2], ignore_index=True) # Combine for full list
data = data[~invalid_mask] # Filter the data



# Calculate participation marks
email_counts = data.groupby(["Email Address", "Section"]).size().reset_index(name="Count")
email_counts["Participation Mark"] = email_counts["Count"] * 0.5
email_counts.loc[email_counts["Section"] == "Section A04, 2-3 pm on Monday and Wednesday", "Participation Mark"] += 0.5
participation_marks_df = email_counts.drop("Count", axis=1)


# Function to process fractional scores
def calculate_score(value, max_points=2):
    try:
        # Split the fraction and calculate the ratio
        numerator, denominator = map(float, value.split("/"))
        return (numerator / denominator) * max_points
    except:
        # Return NaN if value is not in "numerator/denominator" format
        return value


# Clean the "Evaluate" columns
evaluate_columns = ["Evaluate 1", "Evaluate 2", "Evaluate 3", "Evaluate 4", "Evaluate 5"]
for col in evaluate_columns:
    data[col] = data[col].fillna("").astype(str).str.split().str[0] # Get the value before space for "Evaluate" columns
    data[col] = data[col].apply(calculate_score) # Turn fractional scores into standard scores
    data[col] = pd.to_numeric(data[col], errors="coerce")
    data.dropna(subset=[col]) # Remove invalid rows (string or empty)

# Calculate the total presentation score each evaluation gave for each group, and the average score of each group
data["Score"] = data[evaluate_columns].sum(axis=1)
columns_to_mean = evaluate_columns + ["Score"]
presentation_marks_df = data.groupby(["Section", "Topic"])[columns_to_mean].mean().reset_index()

# Save results to files
lied_about_participation_df.to_excel("results/Students_Who_Lied_About_Participation.xlsx")
participation_marks_df.to_excel("results/Participation_Marks.xlsx")
presentation_marks_df.to_excel("results/Presentation_Marks.xlsx")
