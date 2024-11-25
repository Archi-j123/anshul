file_path = r'C:\\Users\\krishankanty\\Desktop\\anshul\\pub_sub.log'

# Open the file in write mode to clear its contents
with open(file_path, 'w') as file:
    # Writing an empty string to the file clears all data
    file.write("")

print("File data pub_sub deleted successfully.")