import pandas as pd
import matplotlib.pyplot as plt

# Specify the file path and sheet name
file_path = "Diversity_data/Diversity_data.xlsx"
sheet_name = "Data"

# List of columns to select
""" Note that there are two rows of column titles. The first row will be skipped to avoid complications with this but this means that some column titles 
    listed below may not make sense on their own. Please refer to the excel sheet for the full column titles"""
# Also note that the “” symbols in some of the column headers must be exactly as they are here or they will not be read correctly i.e. they must be “” and not ""
# This may require editing them in the excel file

columns_to_select = ["What part of the ABL Group do you work for?", "Which best describes your gender identity?", 
                     "Do you describe yourself as trans?", "Which of the following best describes your sexual orientation?",
                     "What is your ethnicity?", "What is your religion?", "How old are you?", "Response", 
                     "Do you experience barriers or limitations in your day-to-day activities related to any disability, health conditions or impairments?",
                     "Response.1", "What type of school did you attend for the majority of your time between the ages of 11 and 16?",
                     "How strongly do you agree with this statement? “I feel like I truly belong at ABL”",
                     "How strongly do you agree with this statement? “I don’t feel like I need to mask or downplay aspects of my physical, cultural, spiritual or emotional self at work”",
                     "How strongly do you agree with this statement? “I believe that ABL is an inclusive employer”",
                     "How strongly do you agree with this statement? “I believe that everyone is able to succeed at ABL, regardless of their background of characteristics”",
                     "How strongly do you agree with this statement? “I feel able to raise equality, diversity or inclusion issues with my line manager or other management at ABL”",
                     "Have you encountered any perceived bias within ABL?Bias being defined a inclination or prejudice for or against one person or group, especially in a way considered to be unfair (Oxford Languages, 2023)"]

# Read the specified columns from the Excel file into a DataFrame, skipping the first row
data = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=1, usecols=columns_to_select)

# Rename some of the columns to be clear as to what they are
column_mapping = {
    "Response": "Do you consider yourself to have a disability or do you have a physical or mental health condition lasting or expected to last 12 months or more?",
    "Response.1": "When you were 18, had any of your parents or guardians completed a university degree course or equivalent (e.g., BA, BSc or higher)?"
}
data.rename(columns=column_mapping, inplace=True)

# Now you can work with the 'data' DataFrame




# Plot a bar chart of group companies
plt.figure(figsize=(10, 7))  # Set the figure size
data['What part of the ABL Group do you work for?'].value_counts().plot(kind='bar')
plt.title('Distribution of Responses Across ABL Group')
plt.xlabel('Group Company')
plt.ylabel('Frequency')
plt.xticks(rotation=45)  # Rotate x-axis labels for better readability
plt.tight_layout()

# Save the plot as a PDF file
output_file = "Diversity_data/figures/Group_Company.pdf"
plt.savefig(output_file, format='pdf')
plt.clf()


# Plot a bar chart of gender
plt.figure(figsize=(10, 7))  # Set the figure size
data['Which best describes your gender identity?'].value_counts().plot(kind='bar')
plt.title('Distribution of Gender')
plt.xlabel('Gender')
plt.ylabel('Frequency')
plt.xticks(rotation=45)  # Rotate x-axis labels for better readability
plt.tight_layout()

# Save the plot as a PDF file
output_file = "Diversity_data/figures/Gender.pdf"
plt.savefig(output_file, format='pdf')
plt.clf()


# Define a custom order for gender
custom_age_order = ['Man', 'Woman', 'Prefer not to say']
# Convert the gender column to Categorical with the custom order
data['Which best describes your gender identity?'] = pd.Categorical(
    data['Which best describes your gender identity?'], 
    categories=custom_age_order,
    ordered=True)

# Group the data by 'What part of the ABL Group do you work for?' and 'Which best describes your gender identity?'
grouped_data = data.groupby(['What part of the ABL Group do you work for?', 'Which best describes your gender identity?']).size().unstack()

# Plot a bar chart
plt.figure(figsize=(10, 7))  # Set the figure size
grouped_data.plot(kind='bar', stacked=False)
plt.title('Gender Identity by Group Company')
plt.xlabel('Group Company')
plt.ylabel('Frequency')
plt.xticks(rotation=45)  # Rotate x-axis labels for better readability
plt.tight_layout()
plt.legend(title='Gender')

# Save the plot as a PDF file
output_file = "Diversity_data/figures/Gender_by_Company.pdf"
plt.savefig(output_file, format='pdf')
plt.clf()


# Plot a bar chart of trans identity
plt.figure(figsize=(10, 7))  # Set the figure size
data['Do you describe yourself as trans?'].value_counts().plot(kind='bar')
plt.title('Distribution of Transgender Identity')
plt.xlabel('Do you describe yourself as trans?')
plt.ylabel('Frequency')
plt.xticks(rotation=45)  # Rotate x-axis labels for better readability
plt.tight_layout()

# Save the plot as a PDF file
output_file = "Diversity_data/figures/Transgender_Identity.pdf"
plt.savefig(output_file, format='pdf')
plt.clf()


# Plot a bar chart of sexual orientation
plt.figure(figsize=(10, 7))  # Set the figure size
data['Which of the following best describes your sexual orientation?'].value_counts().plot(kind='bar')
plt.title('Distribution of Sexual Orientation')
plt.xlabel('Sexual Orientation')
plt.ylabel('Frequency')
plt.xticks(rotation=45)  # Rotate x-axis labels for better readability
plt.tight_layout()

# Save the plot as a PDF file
output_file = "Diversity_data/figures/Sexual_Orientation.pdf"
plt.savefig(output_file, format='pdf')
plt.clf()


# Plot a bar chart of ethnicity
plt.figure(figsize=(10, 8))  # Set the figure size
data['What is your ethnicity?'].value_counts().plot(kind='bar')
plt.title('Distribution of Ethnicity')
plt.xlabel('Ethnicity')
plt.ylabel('Frequency')
# Set the font size of x-axis labels
plt.xticks(rotation=75, fontsize=8)
plt.tight_layout()

# Save the plot as a PDF file
output_file = "Diversity_data/figures/Ethnicity.pdf"
plt.savefig(output_file, format='pdf')
plt.clf()


# Change none entry to be a string so it is registered correctly
data['What is your religion?'].fillna("None", inplace=True)
# Plot a bar chart of religion
plt.figure(figsize=(10, 7))  # Set the figure size
data['What is your religion?'].value_counts().plot(kind='bar')
plt.title('Distribution of Religion')
plt.xlabel('Religion')
plt.ylabel('Frequency')
plt.xticks(rotation=60)
plt.tight_layout()

# Save the plot as a PDF file
output_file = "Diversity_data/figures/Religion.pdf"
plt.savefig(output_file, format='pdf')
plt.clf()


# Define a custom order for age groups
custom_age_order = ['18-24', '25-34', '35-44', '45-54', '55-64', '65+', "Prefer not to say"]
# Convert the age column to Categorical with the custom order
data['How old are you?'] = pd.Categorical(
    data['How old are you?'], 
    categories=custom_age_order,
    ordered=True)
# Group the data by age and count the responses
age_counts = data['How old are you?'].value_counts()

# Plot a bar chart of age
plt.figure(figsize=(10, 7))  # Set the figure size
age_counts.sort_index().plot(kind='bar')
plt.title('Distribution of Age Groups')
plt.xlabel('Age Group')
plt.ylabel('Frequency')
plt.xticks(rotation=45)  # Rotate x-axis labels for better readability
plt.tight_layout()

# Save the plot as a PDF file
output_file = "Diversity_data/figures/Age.pdf"
plt.savefig(output_file, format='pdf')
plt.clf()


# Plot a bar chart of school type
plt.figure(figsize=(10, 7))  # Set the figure size
data['What type of school did you attend for the majority of your time between the ages of 11 and 16?'].value_counts().plot(kind='bar')
plt.title('Distribution of Education')
plt.xlabel('Education from age 11-16')
plt.ylabel('Frequency')
plt.xticks(rotation=45)  # Rotate x-axis labels for better readability
plt.tight_layout()

# Save the plot as a PDF file
output_file = "Diversity_data/figures/Education.pdf"
plt.savefig(output_file, format='pdf')
plt.clf()


# Filter rows with gender identity as "Man" or "Woman"
filtered_data = data[data['Which best describes your gender identity?'].isin(['Man', 'Woman'])]

# Define a custom order
custom_order = ['Strongly disagree', 'Disagree', 'Neither agree nor disagree', 'Agree', 'Strongly agree']
# Convert the response column to Categorical with the custom order
filtered_data['How strongly do you agree with this statement? “I feel like I truly belong at ABL”'] = pd.Categorical(
    filtered_data['How strongly do you agree with this statement? “I feel like I truly belong at ABL”'], 
    categories=custom_order,
    ordered=True)

# Group the filtered data by gender and statement response, and count the occurrences
# Remove rows with blank responses
filtered_data_2 = filtered_data.dropna(subset=['How strongly do you agree with this statement? “I feel like I truly belong at ABL”'])
grouped_data = filtered_data_2.groupby(['How strongly do you agree with this statement? “I feel like I truly belong at ABL”', 'Which best describes your gender identity?']).size().unstack()
# Remove the last column
grouped_data = grouped_data.iloc[:,:2]
# Count the number of people who responded "Man" and "Woman"
man_count = grouped_data['Man'].sum()
woman_count = grouped_data['Woman'].sum()
# Convert the values in the columns to percentages
grouped_data['Man'] = grouped_data['Man'] / man_count * 100
grouped_data['Woman'] = grouped_data['Woman'] / woman_count * 100

# Plot a bar chart
plt.figure(figsize=(10,8))  # Set the figure size
grouped_data.plot(kind='bar', stacked=False)
plt.title('“I feel like I truly belong at ABL”', fontsize=12)
plt.xlabel('Response')
plt.ylabel('Percentage')
plt.xticks(rotation=45, fontsize=8)
plt.tight_layout()
# Remove the legend title
plt.legend(title='')
# Adjust the margins for better vertical spacing
plt.subplots_adjust(bottom=0.3)  # Increase this value to add more space at the bottom

# Save the plot as a PDF file
output_file = "Diversity_data/figures/Belonging_at_ABL_by_Gender.pdf"
plt.savefig(output_file, format='pdf')
plt.clf()


# Convert the response column to Categorical with the custom order
filtered_data['How strongly do you agree with this statement? “I don’t feel like I need to mask or downplay aspects of my physical, cultural, spiritual or emotional self at work”'] = pd.Categorical(
    filtered_data['How strongly do you agree with this statement? “I don’t feel like I need to mask or downplay aspects of my physical, cultural, spiritual or emotional self at work”'], 
    categories=custom_order,
    ordered=True)

# Group the filtered data by gender and statement response, and count the occurrences
# Remove rows with blank responses
filtered_data_2 = filtered_data.dropna(subset=['How strongly do you agree with this statement? “I don’t feel like I need to mask or downplay aspects of my physical, cultural, spiritual or emotional self at work”'])
grouped_data = filtered_data_2.groupby(['How strongly do you agree with this statement? “I don’t feel like I need to mask or downplay aspects of my physical, cultural, spiritual or emotional self at work”', 'Which best describes your gender identity?']).size().unstack()
# Remove the last column
grouped_data = grouped_data.iloc[:,:2]
# Count the number of people who responded "Man" and "Woman"
man_count = grouped_data['Man'].sum()
woman_count = grouped_data['Woman'].sum()
# Convert the values in the columns to percentages
grouped_data['Man'] = grouped_data['Man'] / man_count * 100
grouped_data['Woman'] = grouped_data['Woman'] / woman_count * 100

# Plot a bar chart
plt.figure(figsize=(10,8))  # Set the figure size
grouped_data.plot(kind='bar', stacked=False)
plt.title('“I don’t feel like I need to mask or downplay aspects of my physical, cultural, spiritual or emotional self at work”', fontsize=7.5)
plt.xlabel('Response')
plt.ylabel('Percentage')
plt.xticks(rotation=45, fontsize=8)
plt.tight_layout()
# Remove the legend title
plt.legend(title='')
# Adjust the margins for better vertical spacing
plt.subplots_adjust(bottom=0.3)  # Increase this value to add more space at the bottom

# Save the plot as a PDF file
output_file = "Diversity_data/figures/Downplay_Self_by_Gender.pdf"
plt.savefig(output_file, format='pdf')
plt.clf()


# Convert the response column to Categorical with the custom order
filtered_data['How strongly do you agree with this statement? “I feel able to raise equality, diversity or inclusion issues with my line manager or other management at ABL”'] = pd.Categorical(
    filtered_data['How strongly do you agree with this statement? “I feel able to raise equality, diversity or inclusion issues with my line manager or other management at ABL”'], 
    categories=custom_order,
    ordered=True)

# Group the filtered data by gender and statement response, and count the occurrences
# Remove rows with blank responses
filtered_data_2 = filtered_data.dropna(subset=['How strongly do you agree with this statement? “I feel able to raise equality, diversity or inclusion issues with my line manager or other management at ABL”'])
grouped_data = filtered_data_2.groupby(['How strongly do you agree with this statement? “I feel able to raise equality, diversity or inclusion issues with my line manager or other management at ABL”', 'Which best describes your gender identity?']).size().unstack()
# Remove the last column
grouped_data = grouped_data.iloc[:,:2]
# Count the number of people who responded "Man" and "Woman"
man_count = grouped_data['Man'].sum()
woman_count = grouped_data['Woman'].sum()
# Convert the values in the columns to percentages
grouped_data['Man'] = grouped_data['Man'] / man_count * 100
grouped_data['Woman'] = grouped_data['Woman'] / woman_count * 100
# Plot a bar chart
plt.figure(figsize=(10,8))  # Set the figure size
grouped_data.plot(kind='bar', stacked=False)
plt.title('“I feel able to raise equality, diversity or inclusion issues with my line manager or other management at ABL”', fontsize=7.5)
plt.xlabel('Response')
plt.ylabel('Percentage')
plt.xticks(rotation=45, fontsize=8)
plt.tight_layout()
# Remove the legend title
plt.legend(title='')
# Adjust the margins for better vertical spacing
plt.subplots_adjust(bottom=0.3)  # Increase this value to add more space at the bottom

# Save the plot as a PDF file
output_file = "Diversity_data/figures/Able_to_Raise_Equality_by_Gender.pdf"
plt.savefig(output_file, format='pdf')
plt.clf()


# Group the filtered data by gender and statement response, and count the occurrences
# Remove rows with blank responses
filtered_data_2 = filtered_data.dropna(subset=['Have you encountered any perceived bias within ABL?Bias being defined a inclination or prejudice for or against one person or group, especially in a way considered to be unfair (Oxford Languages, 2023)'])
grouped_data = filtered_data_2.groupby(['Have you encountered any perceived bias within ABL?Bias being defined a inclination or prejudice for or against one person or group, especially in a way considered to be unfair (Oxford Languages, 2023)', 'Which best describes your gender identity?']).size().unstack()
# Remove the last column
grouped_data = grouped_data.iloc[:,:2]
# Count the number of people who responded "Man" and "Woman"
man_count = grouped_data['Man'].sum()
woman_count = grouped_data['Woman'].sum()
# Convert the values in the columns to percentages
grouped_data['Man'] = grouped_data['Man'] / man_count * 100
grouped_data['Woman'] = grouped_data['Woman'] / woman_count * 100

# Plot a bar chart
plt.figure(figsize=(10,8))  # Set the figure size
grouped_data.plot(kind='bar', stacked=False)
plt.title('Have you encountered any perceived bias within ABL?', fontsize=10)
plt.xlabel('Response')
plt.ylabel('Percentage')
plt.xticks(rotation=45, fontsize=12)
plt.tight_layout()
# Remove the legend title
plt.legend(title='')
# Adjust the margins for better vertical spacing
plt.subplots_adjust(bottom=0.3)  # Increase this value to add more space at the bottom

# Save the plot as a PDF file
output_file = "Diversity_data/figures/Perceived_Bias_by_Gender.pdf"
plt.savefig(output_file, format='pdf')
plt.clf()