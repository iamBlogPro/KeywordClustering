# Keyword Clustering v1.0

Takes a CSV file of keywords - then harnesses Topic Modeling using BerTopic to group them into relevant clusters. Clusters are then stored in spreadsheets, along with outliers. 

# Script for automated keyword clustering using BERTopic
BERTopic is a topic modeling technique that leverages transformer models and c-TF-IDF to efficiently cluster text data into topics. This script provides a streamlined way to preprocess a collection of keywords, cluster them into topics, save the results to Excel, and optionally visualize the topic distribution.

## Requirements

- Python 3.6 or higher
- Install required packages: `pip install -r requirements.txt`

## Usage 

1. Place your list of keywords in a CSV file with no header (e.g., `keywords.csv`).

2. Optionally, edit `config.json` and adjust parameters
   
3. Run the script:

```bash
python cluster.py
```

4. The script will generate two output files:
    - `clustered_keywords.xlsx`: An Excel spreadsheet with color-coded groups of your keywords. 
    - `outliers.xlsx`: An Excel file with keywords, that couldn't be grouped - your outliers, basically.
  
## How does it work?

1. **Reads keywords from a CSV file** - The script reads your list of keywords from a CSV file, storing the keywords in a list.
2. **Initialization** - The script sets up logging and loads the configuration from a JSON file.
3. **Preprocessing** - The keywords are preprocessed (converted to lowercase and stripped of leading/trailing spaces).
4. **Topic Modeling** - BERTopic is used to perform topic modeling on the preprocessed keywords.
5. **Result Saved** - The topics and outlier keywords are saved into separate Excel files, with visual formatting for easy distinction of topics.
6. **Visualization and Model Saving (Optional)** - If specified in the configuration, the script will visualize the topic clusters using a seaborn-generated color palette and save the BERTopic model with a timestamp for future use.

## Highlights

### **BERTopic Integration**
Utilizes the BERTopic model for advanced topic modeling, which is based on state-of-the-art transformer architectures.
### **Configurable**
Allows custom configurations through a JSON file to tailor the clustering process to specific needs.
### **Data Preprocessing**
Includes preprocessing steps to clean keyword data for optimal clustering performance.
### **Logging**
Features a robust logging mechanism to track the process and capture any issues during execution.
### **Visualization**
Offers an option to visualize the topic clusters when the script is configured to do so.
### **Excel Output**
The clustered keywords along with their assigned topics are saved into Excel format for easy analysis and reporting.
### **Separate Outlier Handling**
Outliers, or keywords not fitting well into any topic, are saved into a separate Excel file.
### **Model Saving**
Provides functionality to save the trained BERTopic model for future use or incremental training.

