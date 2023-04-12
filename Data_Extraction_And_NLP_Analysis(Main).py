import Stop_Word
print('Stop Word File Created')
import nltk
nltk.download('punkt')
from nltk.tokenize import sent_tokenize,word_tokenize
from nltk.corpus import stopwords
from bs4 import BeautifulSoup
import string
import requests, openpyxl
i=0
#Counts  number of syllables in a given word.
def count_syllables(word):
    
    vowels = 'aeiouy'
    num_vowels = 0
    last_was_vowel = False
    for char in word:
        if char in vowels:
            if not last_was_vowel:
                num_vowels += 1
            last_was_vowel = True
        else:
            last_was_vowel = False

    # Handle some exceptions
    if word[-2:] == 'es' or word[-2:] == 'ed':
        num_vowels -= 1
    if word[-1:] == 'e' and num_vowels > 1:
        num_vowels -= 1

    return num_vowels
#create a  dictionary for text filteration
with open('positive-words.txt',mode='r', encoding='cp1252') as n:
    lines_n=[line.strip() for line in n.readlines()]
with open('Negative-words.txt',mode='r', encoding='cp1252') as p:
    lines_p=[line.strip() for line in p.readlines()]
key1='Positive'
key2='Negative'
dictt={key1:lines_n,key2:lines_p}

#Get the input url from the given xlsx file to ctreate a list
workbook= openpyxl.load_workbook('Input.xlsx')
worksheet= workbook['Sheet1']
Column_Values=[]
i=0
t=[]
for row in worksheet.iter_rows(min_row=2,min_col=2,values_only=True):
    Column_Values.append(row[0])
print('Excel sheets all URL links are Extracted')
#get the stop words from the txt file 
t=[]
with open('Master1.txt',mode='r', encoding='cp1252') as stri:
    stop_words=[line.strip() for line in stri.readlines()]
#print(stop_words)
print(f'===================================================={i}==========================================')
#Assign values to the Xlsx file cell index
row_index = 2 
col_index = 3
#Now Start the loop   with arg as list value'Column_Values'
for link in Column_Values:
  response=requests.get(link)
  print('Readed')
  soup=BeautifulSoup(response.text,'html.parser')
  #print(soup)
  articl=soup.find("h1",class_='entry-title')
  print(articl)
  # if - Condition is created for checking the page is valied or Not
  if articl is not None:
    article=articl.get_text(strip=True).replace(string.punctuation,'')
    print("Heading Extracted")
    t.append(article) 

    txt= soup.find("div",class_='td-ss-main-content')
    tags=txt.find_all(['p','h3'])
        
    for tag in tags:
     cont=tag.get_text()
     t.append(cont)
# The Section of Text Analysis is started 
    token=' '.join(t)
    tokns=sent_tokenize(token)
    number_Of_Words_In_Sentence=[]
    for sent in tokns:
      lenth=len(word_tokenize(sent))
      number_Of_Words_In_Sentence.append(lenth)
    Average_Words_per_sentence=sum(number_Of_Words_In_Sentence)/len(number_Of_Words_In_Sentence)
    Total_no_Sentence=len(tokns)
    tokens=word_tokenize(str(tokns))
    filtered1_tokens = [token for token in tokens if token.lower() and token.upper() not in stop_words]
    #print(filtered1_tokens)
    filtered2_tokens =[a for a in filtered1_tokens if a not in string.punctuation]
    Total_no_Words=len(filtered2_tokens)
    syllable_counts = []
    Total_num_Of_Char_IN_Word=0
    for word in filtered2_tokens:
      syllables = count_syllables(word)
      syllable_counts.append(syllables)
      for s in word:
       Total_num_Of_Char_IN_Word+=1
    complex_n=0
    for count in syllable_counts:
      if count >2:
        complex_n+=1
    complex_Word_count=complex_n
    print(filtered2_tokens)
    print(len(filtered2_tokens))
    count_positive = 0
    count_negative = 0
# Iterate through the list and check if each value is in the dictionary
    for value in filtered2_tokens:
      if value in dictt[key1]:
        count_positive += 1
      if value in dictt[key2]:
        count_negative -= 1
    Personal_Noun='--'
    count_polarity=(count_positive-count_negative)/((count_positive+count_negative)+0.000001)
    Subjective_Score=(count_positive+count_negative)/((Total_no_Words)+0.000001)
    Average_Sentence_length = Total_no_Words/Total_no_Sentence
    Percentage_Of_Complex=complex_Word_count/Total_no_Words
    Fog_Index=0.4*(Average_Sentence_length+Percentage_Of_Complex)
    Avg_Word_Length=Total_num_Of_Char_IN_Word/Total_no_Words
    Number_Of_Words_In_Sentence=sum(number_Of_Words_In_Sentence)
    syllable_counts=sum(syllable_counts)
# Print the counts for each key
  
    #print("Count for Positive:", count_positive)
    #print("Count for Negative:", count_negative)
    #print(Subjective_Score)
   # print(count_polarity)
   # print(Average_Sentence_length)
   # print(Percentage_Of_Complex)
   # print(complex_Word_count)
   # print(Fog_Index)
   # print(Average_Words_per_sentence)
   # print(Number_Of_Words_In_Sentence)
   # print(Total_no_Words)
   # print(len(syllable_counts))
   # print(Avg_Word_Length)
#==========================================================================================================================
#write Values to the xlsx file
    values= [count_positive,count_negative,count_polarity,Subjective_Score,Average_Sentence_length,Percentage_Of_Complex,Fog_Index,Average_Words_per_sentence,complex_Word_count,Total_no_Words,syllable_counts,Personal_Noun,Avg_Word_Length]
    wb = openpyxl.load_workbook('Output Data Structure.xlsx')
    ws = wb.active
    for val in values:
      val=str(val)
      ws.cell(row=row_index, column=col_index, value=val.strip())
      col_index += 1
      if col_index > 15:
          row_index += 1
          col_index = 3
    wb.save("Output Data Structure.xlsx")
    print('Done!!!')
    i+=1
    print(f'===================================================={i}==========================================')
  else:# if the Web page is not have Content value as 'None'
    count_positive='None'
    count_negative='None'
    Subjective_Score='None'
    count_polarity='None'
    Average_Sentence_length='None'
    Percentage_Of_Complex='None'
    Fog_Index='None'
    complex_Word_count='None'
    Average_Words_per_sentence='None'
    number_Of_Words_In_Sentence='None'
    Total_no_Words='None'
    syllable_counts='None'
    Personal_Noun='None'
    #print(count_positive)
    #print(count_negative)
   # print(Subjective_Score)
   # print(count_polarity)
   # print(Average_Sentence_length)
   # print(Percentage_Of_Complex)
   # print(Fog_Index)
   # print(Average_Words_per_sentence)
   # print(number_Of_Words_In_Sentence)
   # print(Total_no_Words)
   # print(syllable_counts)
    values = [count_positive,count_negative,count_polarity,Subjective_Score,Average_Sentence_length,Percentage_Of_Complex,Fog_Index,Average_Words_per_sentence,complex_Word_count,Total_no_Words,syllable_counts,Personal_Noun,Avg_Word_Length]
    wb = openpyxl.load_workbook('Output Data Structure.xlsx')
    ws = wb.active
    for val in values:
      val=str(val)
      ws.cell(row=row_index, column=col_index, value=val.strip())
      col_index += 1
      if col_index > 15:
          row_index += 1
          col_index = 3
    wb.save("Output Data Structure.xlsx")
print('Done!!!!!!!!!!!!!!!!!!!!!!')
