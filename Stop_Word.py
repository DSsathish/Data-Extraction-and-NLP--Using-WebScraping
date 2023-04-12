import os
stop_words = set()

# Define the directory where the stopword files are located
dir_path = "C:/Users/kings/Desktop/BlackCoffer_Intern/StopWords/"

# Loop through each file in the directory
for filename in os.listdir(dir_path):
    # Read the file and append its words to the set of stopwords
    with open(os.path.join(dir_path, filename), 'r') as file:
        stop_words.update(word.strip() for word in file.readlines())
with open('Master1.txt', 'w') as file:
    for n in stop_words:
        file.write('%s\n' % n)
    file.close()
    
print('File is cteated!')