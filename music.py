import os
import pandas as pd

# ---------------------------------------------------Analyze txt file---------------------------------------------------
music_data = []  # Create a list to store dictionaries

with open('C:\\Users\\Grigoris\\Desktop\\music.txt', 'r', encoding='utf-8') as music_file:
    for row in music_file.readlines():
        music_dict = {}
        music_ls = [item.strip() for item in row.split("\t")]
        # print(music_ls)
        music_dict['Track'] = music_ls[0]
        music_dict['Time'] = music_ls[2]
        music_dict['Artist'] = music_ls[3]
        music_dict['Genre'] = music_ls[5]
        music_dict['played'] = music_ls[6]
        music_dict["Year"] = music_ls[8]
        music_data.append(music_dict)

## Print the resulting list of dictionaries
# for entry in music_data:
#     print(entry)



# ---------------------------------------------------Analyze Artists---------------------------------------------------
artist_dict = {}
for entry in music_data:
    artist = entry["Artist"]
    artist_dict[artist] = artist_dict.get(artist,0)+1

sorted_artist_dict = dict(sorted(artist_dict.items(), key=lambda item: item[1], reverse=True))

## Print the sorted dictionary
# for k,v in sorted_artist_dict.items():
#     print(k, v)


# ---------------------------------------------------Analyze Genre---------------------------------------------------
genre_dict = {}
for entry in music_data:
    genre = entry["Genre"]
    genre_dict[genre] = genre_dict.get(genre,0)+1

sorted_genre_dict = dict(sorted(genre_dict.items(), key=lambda item: item[1], reverse=True))

# # Print the sorted dictionary
# for k,v in sorted_genre_dict.items():
#     print(k, v)




# ---------------------------------------------------Create Music.xlsx file---------------------------------------------------

# Convert the songs dictionary to a DataFrame
df_songs = pd.DataFrame((music_data), columns=['Track','Artist', 'Time', 'Year', 'Genre'])

# Convert the artist dictionary to a DataFrame
df_artist = pd.DataFrame(list(sorted_artist_dict.items()), columns=['Artist', 'Count'])

# Convert the genre dictionary to a DataFrame
df_genre = pd.DataFrame(list(sorted_genre_dict.items()), columns=['Genre', 'Count'])


# Save the DataFrame to an Excel file with a specific sheet name and autofit column widths
with pd.ExcelWriter('C:\\Users\\Grigoris\\Desktop\\Music.xlsx', engine='xlsxwriter') as writer:

    # ---------------------------------------------------Export Tracks---------------------------------------------------

    df_songs[['Track','Artist', 'Time', 'Year', 'Genre']].to_excel(writer, sheet_name='Tracks', index=False)

    # Access the XlsxWriter workbook and worksheet objects from the DataFrame writer object
    workbook = writer.book
    worksheet = writer.sheets['Tracks']
    worksheet.write_string('F1', 'Number of tracks')
    worksheet.write_formula("F2", 'COUNTA(A:A)-1')

    # Autofit the columns in the 'Tracks' sheet
    for i, col in enumerate(df_songs.columns):
        max_len = max(df_songs[col].astype(str).apply(len).max(), len(col)) + 2
        worksheet.set_column(i, i, max_len)




    # ---------------------------------------------------Export Artists---------------------------------------------------

    df_artist.to_excel(writer, sheet_name='Artist', index=False)

    # Access the XlsxWriter workbook and worksheet objects from the DataFrame writer object
    workbook = writer.book
    worksheet = writer.sheets['Artist']

    # Autofit the columns in the 'Artist' sheet
    for i, col in enumerate(df_artist.columns):
        max_len = max(df_artist[col].astype(str).apply(len).max(), len(col)) + 2
        worksheet.set_column(i, i, max_len)


    # Create a bar chart with the top 10 artists
    chart = workbook.add_chart({'type': 'bar'})

    # Configure the chart data
    chart.add_series({
        'name': ['Artist', 0, 0],
        'categories': ['Artist', 1, 0, 11, 0],  # Assuming your data starts from row 1
        'values': ['Artist', 1, 1, 11, 1],
        'data_labels': {'value': True,},

    })
    chart.set_y_axis({'reverse': True})

    chart.set_size({'x_scale': 1.5, 'y_scale': 2})
    chart.set_legend({'none': True})


    # Insert the chart into the 'Artist' sheet starting at cell G5
    worksheet.insert_chart('G5', chart)


    # ---------------------------------------------------Export Genre---------------------------------------------------

    df_genre.to_excel(writer, sheet_name='Genre', index=False)

    # Access the XlsxWriter workbook and worksheet objects from the DataFrame writer object
    workbook = writer.book
    worksheet = writer.sheets['Genre']

    # Autofit the columns in the 'Artist' sheet
    for i, col in enumerate(df_artist.columns):
        max_len = max(df_artist[col].astype(str).apply(len).max(), len(col)) + 2
        worksheet.set_column(i, i, max_len)

# Create a bar chart with the top 10 genre
    chart = workbook.add_chart({'type': 'bar'})

    # Configure the chart data
    chart.add_series({
        'name': ['Genre', 0, 0],
        'categories': ['Genre', 1, 0, 11, 0],  # Assuming your data starts from row 1
        'values': ['Genre', 1, 1, 11, 1],
        'data_labels': {'value': True,},

    })
    chart.set_y_axis({'reverse': True})

    chart.set_size({'x_scale': 1.5, 'y_scale': 2})
    chart.set_legend({'none': True})


    # Insert the chart into the 'Genre' sheet starting at cell G5
    worksheet.insert_chart('G5', chart)





# ---------------------------------------------------Open xlsx file---------------------------------------------------

excel_file_path = 'C:\\Users\\Grigoris\\Desktop\\Music.xlsx'  # Change this to your actual file path
try:
    os.system(f'start excel {excel_file_path}')  # Adjust this line based on your OS and Excel path
except Exception as e:
    print(f"Unable to open Excel file: {e}")

