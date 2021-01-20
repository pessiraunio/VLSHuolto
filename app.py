import folium
import pandas as pd
#from converter import Prettify



#Kerätään data excelistä.

data_file = pd.ExcelFile(r"C:\\Users\\Pepo\\Documents\\Python\\VLS-Huolto\\VLSHuolto\\excel.xlsx") #Polku pöytäkoneella
#data_file = pd.ExcelFile(r"C:\\Users\\pessi\Documents\\Python Scripts\\VLS-Huollot\\excel.xlsx") #Polku läppärillä

columns_in_file = pd.read_excel(data_file,"Sheet1")

class MainClass():
    #Tehdään jokaisesta sarakkeesta oma lista.
    Merkki = columns_in_file['Merkki'].tolist()
    Malli = columns_in_file['Malli'].tolist()
    Huollettu = columns_in_file['Huollettu'].tolist()
    Sijainti = columns_in_file['Sijainti'].tolist()
    #sijaintitiedot
    Northern = pd.read_excel(data_file, "Sheet1", usecols="C")
    Eastern = pd.read_excel(data_file, "Sheet1", usecols="D")


    #rakennetaan kartta
    map_fin = folium.Map(location=[65.369781, 27.171806], zoom_start=6, width='60%', height='100%')

    #Html ja iframen testailua
    html = f'''<p><b>Merkki</b></p><br>{Merkki[1]}
    <p><b>Malli</b></p> <br>{Malli[1]}
    <p><b>Huollettu</b></p> <br>{Huollettu[1]}'''

    frame = folium.IFrame(html, width=100, height=100)

    popup = folium.Popup(frame,max_width=100)

    testmarker = folium.Marker(location=[65.369781, 27.171806], popup=popup).add_to(map_fin)

    #For loop rivien selaamiseen
    column_num = 0
    for column in columns_in_file:
        column_num = column_num + 1

    for i in range(column_num):
        tooltip=("Klikkaa lisätietoja") #Tiedot kun leijuu hiirellä
        folium.map.Marker(
            location=[Northern.loc[i], Eastern.loc[i]], 
            popup=(f'''<p><b>Merkki</b></p><br>{Merkki[i]}
                        <p><b>Malli</b></p> <br>{Malli[i]}
                        <p><b>Huollettu</b></p> <br>{Huollettu[i]}'''), 
            tooltip=tooltip).add_to(map_fin)

    #luodaan html.sivu
    map_fin.save('site.html')


    print(Huollettu)

