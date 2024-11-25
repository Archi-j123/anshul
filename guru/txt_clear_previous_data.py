file_path = r'C:\\Users\\krishankanty\\Desktop\\anshul\\data_guruX.txt'

# Open the file in write mode to clear its contents
with open(file_path, 'w') as file:
    # Writing an empty string to the file clears all data
    file.write("")

print("File text data deleted successfully.")
