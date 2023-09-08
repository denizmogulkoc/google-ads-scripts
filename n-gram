import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
import string
import nltk
from nltk.corpus import stopwords
from collections import defaultdict

# Load CSV data
df = pd.read_csv('location.csv') # location 

# Extract features and target
y = df["Sentiment"].values
x = df["Description"].values

# Split data into training and testing sets
(x_train, x_test, y_train, y_test) = train_test_split(x, y, test_size=0.1)

# Create DataFrames for training and testing data
df_train = pd.DataFrame({'description': x_train, 'sentiment': y_train})
df_test = pd.DataFrame({'description': x_test, 'sentiment': y_test})

# Define function to remove punctuation
def remove_punctuation(text):
    if isinstance(text, float):
        return text
    ans = ""
    for i in text:
        if i not in string.punctuation:
            ans += i
    return ans

# Apply remove_punctuation function to description column
df_train['description'] = df_train['description'].apply(remove_punctuation)
df_test['description'] = df_test['description'].apply(remove_punctuation)

# Download NLTK stopwords
nltk.download('stopwords')

# Define function to generate N-grams
def generate_N_grams(text, ngram=1):
    if isinstance(text, float):
        return []  # Return an empty list if text is NaN
    words = [word for word in text.split(" ") if word not in set(stopwords.words('english'))]
    temp = zip(*[words[i:] for i in range(0, ngram)])
    ans = [' '.join(ngram) for ngram in temp]
    return ans

# Initialize defaultdicts for sentiment values
promisingValues = defaultdict(int)
coreValues = defaultdict(int)
newValues = defaultdict(int)

# Iterate through descriptions based on sentiment
for sentiment, description in zip(df_train['sentiment'], df_train['description']):
    for word in generate_N_grams(description):
        if sentiment == "Promising":
            promisingValues[word] += 1
        elif sentiment == "Core":
            coreValues[word] += 1
        elif sentiment == "New":
            newValues[word] += 1

# Get the top 10 recurring words for each sentiment
top_promising_words = sorted(promisingValues.items(), key=lambda x: x[1], reverse=True)[:10]
top_core_words = sorted(coreValues.items(), key=lambda x: x[1], reverse=True)[:10]
top_new_words = sorted(newValues.items(), key=lambda x: x[1], reverse=True)[:10]

# Separate words and counts for plotting
pro1, pro2 = zip(*top_promising_words)
core1, core2 = zip(*top_core_words)
new1, new2 = zip(*top_new_words)

# Define function to plot top words
def plot_top_words(words, counts, title):
    plt.figure(figsize=(16, 4))
    plt.bar(words, counts, color='green', width=0.4)
    plt.xlabel("Words")
    plt.ylabel("Count")
    plt.title(title)
    plt.xticks(rotation=45)
    plt.show()

# Plot top words for each sentiment
plot_top_words(pro1, pro2, "Top 10 words in Promising sentiment")
plot_top_words(core1, core2, "Top 10 words in Core sentiment")
plot_top_words(new1, new2, "Top 10 words in New sentiment")
