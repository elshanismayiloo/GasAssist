import pandas as pd
import webbrowser
import kivy
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.scrollview import ScrollView

# Load Excel File
excel_file_path = 'database1.xlsx'
df = pd.read_excel(excel_file_path)

class SubscriberApp(App):
    def build(self):
        self.title = 'Subscriber Information'
        layout = BoxLayout(orientation='vertical', spacing=10, padding=10)
        
        self.search_input = TextInput(hint_text='Enter GmNo, SubNo, or Fullname')
        search_button = Button(text='Search', on_press=self.search)
        self.info_label = Label(text='', halign='left', markup=True)
        self.scrollview = ScrollView(do_scroll_x=False)
        
        layout.add_widget(self.search_input)
        layout.add_widget(search_button)
        layout.add_widget(self.scrollview)
        layout.add_widget(self.info_label)
        
        return layout
    
    def search(self, instance):
        search_value = self.search_input.text
        filtered_data = df[df.apply(lambda row: search_value in [str(row['GmNo']), str(row['SubNo']), str(row['Fullname'])], axis=1)]

        if not filtered_data.empty:
            content = BoxLayout(orientation='vertical', spacing=5, size_hint_y=None)
            content.bind(minimum_height=content.setter('height'))

            for _, row in filtered_data.iterrows():
                address = row['Address']
                lat_lon = row['LatitudeLongitude']

                info = f"[b]Fullname:[/b] {row['Fullname']}\n"
                info += f"[b]SubNo:[/b] {row['SubNo']}\n"
                info += f"[b]GmNo:[/b] {row['GmNo']}\n"
                info += f"[b]Address:[/b] {address}\n"

                if 'http' in address:
                    info += f"[ref={address}]Open Map[/ref]"

                label = Label(text=info, markup=True, size_hint_y=None)
                label.bind(on_ref_press=lambda _, url: self.open_map(address, lat_lon))
                content.add_widget(label)
            
            self.scrollview.clear_widgets()
            self.scrollview.add_widget(content)
            self.info_label.text = ''
        else:
            self.info_label.text = 'No matching data found for the entered value.'
            self.scrollview.clear_widgets()
    
    def open_map(self, address, lat_lon):
        lat, lon = map(float, lat_lon.split(','))
        m = folium.Map(location=[lat, lon], zoom_start=15)
        folium.Marker([lat, lon], popup=folium.Popup(address, parse_html=True)).add_to(m)
        m.save('map.html')
        webbrowser.open('map.html')

if __name__ == '__main__':
    SubscriberApp().run()
