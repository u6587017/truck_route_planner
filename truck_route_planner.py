import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
import pandas as pd
import folium
import webbrowser
from openpyxl import Workbook
from folium import plugins
import datetime
import numpy as np
from numpy import sin,cos,sqrt,arctan2,radians

def main():
    # Sample depot location
    depot_location = (14.0828151, 100.6258423)  # Assuming the depot location (Bangkok, Thailand)

    # Global variables to hold the full DataFrame and selected date
    full_df = None
    selected_date = None

    # Function to read and clean the Excel file
    def read_excel(file_path):
        df = pd.read_excel(file_path)
        df['Day of Completed At Order'] = pd.to_datetime(df['Day of Completed At Order'], format='%d-%m-%Y')
        df = df[['Day of Completed At Order', 'Order Number', 'Lat', 'Lng', 'Total weights per order', 'Region', 'Address Province Group', 'Ship Province', 'Geo District']]
        orders = []
        for i in range(len(df)):
            orders.append({
                'id': df.at[i, 'Order Number'],
                'weight': df.at[i, 'Total weights per order'],
                'dst': (df.at[i, 'Lat'], df.at[i, 'Lng']),
                'date': df.at[i, 'Day of Completed At Order'],
                'region': df.at[i, 'Region'],
                'address_province_group': df.at[i, 'Address Province Group'],
                'ship_province': df.at[i, 'Ship Province'],
                'geo_district': df.at[i, 'Geo District']
            })
        return orders, df

    # Haversine formula to calculate the distance between two lat/lng pairs
    def haversine(coord1, coord2):
        R = 6371.0  # Earth radius in kilometers

        lat1, lon1 = map(radians, coord1)
        lat2, lon2 = map(radians, coord2)

        dlon = lon2 - lon1
        dlat = lat2 - lat1

        a = sin(dlat / 2)**2 + cos(lat1) * cos(lat2) * sin(dlon / 2)**2
        c = 2 * arctan2(sqrt(a), sqrt(1 - a))

        distance = R * c
        return distance

    # Function to find the nearest neighbor
    def find_nearest_neighbor(current_location, remaining_orders):
        nearest_order = None
        min_distance = float('inf')
        
        for order in remaining_orders:
            distance = haversine(current_location, order['dst'])
            if distance < min_distance:
                min_distance = distance
                nearest_order = order
                
        return nearest_order

    # Function to select the set of orders for a truck
    def select_truck_orders(orders, truck_capacity):
        selected_orders = []
        current_weight = 0
        current_location = depot_location

        while orders:
            nearest_order = find_nearest_neighbor(current_location, orders)
            if current_weight + nearest_order['weight'] <= truck_capacity:
                selected_orders.append(nearest_order)
                current_weight += nearest_order['weight']
                orders.remove(nearest_order)
                current_location = nearest_order['dst']
            else:
                break
        
        return selected_orders

    # Function to generate a list of colors
    def generate_colors():
        return [
            '#FF0000',  # red
            '#0000FF',  # blue
            '#808080',  # gray
            '#8B0000',  # darkred
            '#FFCCCB',  # lightred
            '#FFA500',  # orange
            '#F5F5DC',  # beige
            '#006400',  # darkgreen
            '#90EE90',  # lightgreen
            '#00008B',  # darkblue
            '#ADD8E6',  # lightblue
            '#800080',  # purple
            '#301934',  # darkpurple
            '#5F9EA0',  # cadetblue
            '#D3D3D3',  # lightgray
            '#000000'   # black
        ]

    # Function to create and display the map
    def create_map(orders):
        global all_routes  # Make all_routes accessible to other functions
        remaining_orders = orders.copy()
        truck_number = 1
        all_routes = []

        while remaining_orders:
            truck_orders = select_truck_orders(remaining_orders, 350)
            if not truck_orders:
                break
            all_routes.append(truck_orders)
            
            # Calculate total weight of the truck
            total_weight = sum(order['weight'] for order in truck_orders)
            
            # Print order details and total weight for the current truck
            print(f'Orders for Truck {truck_number}:')
            for order in truck_orders:
                print(f"Order ID: {order['id']}, Weight: {order['weight']} kg, Destination: {order['geo_district']}")
            print(f"Total weight of Truck {truck_number}: {total_weight} kg\n")
            
            truck_number += 1

        if not all_routes:
            print("No routes to display.")

        # Create a map centered around the depot location
        m = folium.Map(location=depot_location, zoom_start=12)
        
        # Add depot marker
        folium.Marker(depot_location, popup='Depot', icon=folium.Icon(color='green')).add_to(m)
        
        # Generate a list of colors for the trucks
        colors = generate_colors()
        name_color = ['red', 'blue', 'gray', 'darkred', 'lightred', 'orange', 'beige', 'darkgreen', 'lightgreen', 'darkblue', 'lightblue', 'purple', 'darkpurple', 'cadetblue', 'lightgray', 'black']
        # Sidebar HTML
        sidebar_html = """
        <div id="sidebar" style="position: fixed; top: 50px; left: 10px; width: 300px; height: 500px; 
            background: white; border: 1px solid black; z-index: 1000; overflow-y: auto;">
            <h2>Truck Orders</h2>
            <form>
            {checkboxes}
            </form>
            {content}
        </div>
        <script>
        function toggleRoute(truckNumber) {{
            var route = document.getElementById('route' + truckNumber);
            if (route.style.display === 'none') {{
                route.style.display = 'block';
            }} else {{
                route.style.display = 'none';
            }}
        }}
        </script>
        """

        sidebar_content = ""
        checkboxes = ""

        for i, truck_orders in enumerate(all_routes):
            truck_number = i + 1
            color = colors[i % len(colors)]
            n_color = name_color[i%len(name_color)]
            route = [depot_location] + [order['dst'] for order in truck_orders]
            
            total_orders = len(truck_orders)
            total_weight = sum(order['weight'] for order in truck_orders)
            truck_info = f'<h4 style="color:{color};">Truck {truck_number} Line: {n_color} line<br/> Total weight: {total_weight:.2f} kg<br/> Total orders:{total_orders}</h4>'
            truck_info += '<table style="width:100%; border-collapse: collapse;">'
            truck_info += '<tr><th style="border: 1px solid black;">Order ID</th><th style="border: 1px solid black;">Weight (kg)</th><th style="border: 1px solid black;">Destination</th></tr>'
            j = 1

            for order in truck_orders:
                popup_text = (f"Truck {truck_number} - {order['id']}, "
                            f"Weight: {order['weight']} kg, Destination: {order['geo_district']}")
                folium.Marker(order['dst'], popup=popup_text,
                            icon=plugins.BeautifyIcon(
                            icon="arrow-down", icon_shape="marker",
                            number=str(j),
                            background_color= color
                        )).add_to(m)
                truck_info += f"<tr><td style='border: 1px solid black;'>{order['id']}</td><td style='border: 1px solid black;'>{order['weight']}</td><td style='border: 1px solid black;'>{order['geo_district']}</td></tr>"
                j += 1
            
            truck_info += "</table>"
            sidebar_content += f'<div id="route{truck_number}" style="display:block;">{truck_info}</div>'
            checkboxes += f'<input type="checkbox" onclick="toggleRoute({truck_number})" checked> Truck {truck_number}<br>'

            # Draw the route
            folium.PolyLine(route, color=color, weight=3, opacity=1, id=f'route{truck_number}').add_to(m)

        sidebar_html = sidebar_html.format(checkboxes=checkboxes, content=sidebar_content)
        
        # Add the sidebar HTML to the map
        m.get_root().html.add_child(folium.Element(sidebar_html))
        
        # Save map to an HTML file and display it
        m.save('truck_routes.html')
        webbrowser.open('truck_routes.html')

    # Function to export the orders to an Excel file
    def export_to_excel():
        global full_df
        try:
            wb = Workbook()
            for i, truck_orders in enumerate(all_routes):
                ws = wb.create_sheet(title=f'Truck {i + 1}')
                ws.append(full_df.columns.tolist())
                for order in truck_orders:
                    row_data = full_df.loc[full_df['Order Number'] == order['id']].values.flatten().tolist()
                    ws.append(row_data)
            # Remove the default sheet created
            del wb['Sheet']
            file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[("Excel files", "*.xlsx *.xls")])
            if file_path:
                wb.save(file_path)
                messagebox.showinfo("Success", "The orders have been exported successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while exporting the file:\n{e}")

    # Function to select an Excel file and generate the map
    def select_file():
        global full_df, selected_date
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                orders, full_df = read_excel(file_path)
                if selected_date and selected_date != 'All':
                    selected_date = datetime.datetime.strptime(selected_date, '%d-%m-%Y').strftime('%Y-%m-%d')
                    orders = [order for order in orders if order['date'].strftime('%Y-%m-%d') == selected_date]
                create_map(orders)
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while processing the file:\n{e}")

    # Function to set the selected date
    def set_date(date):
        global selected_date
        selected_date = date

    # Function to show a popup message when "All Dates" button is clicked
    def all_dates_popup():
        messagebox.showinfo("All Dates Selected", "You have selected to include orders from all dates.")
        set_date('All')

    # Tkinter GUI
    root = tk.Tk()
    root.title("Truck Route Planner")

    # Set the background color
    root.configure(bg="#f0f0f0")

    frame = tk.Frame(root, bg="#f0f0f0")
    frame.pack(pady=20)

    label = tk.Label(frame, text="Truck Route Planner", font=("Arial", 24, "bold"), bg="#f0f0f0", fg="#333")
    label.pack(pady=10)

    date_label = tk.Label(frame, text="Select Date:", font=("Arial", 16), bg="#f0f0f0", fg="#333")
    date_label.pack(pady=10)

    date_entry = DateEntry(frame, width=12, font=("Arial", 14), background="darkblue", foreground="white", borderwidth=2, date_pattern='dd-MM-yyyy')
    date_entry.pack(pady=10)
    date_entry.bind("<<DateEntrySelected>>", lambda event: set_date(date_entry.get()))

    all_dates_button = tk.Button(frame, text="All Dates", command=all_dates_popup, font=("Arial", 16, "bold"), bg="#4CAF50", fg="white", padx=10, pady=5)
    all_dates_button.pack(pady=10)

    select_button = tk.Button(frame, text="Select Excel File", command=select_file, font=("Arial", 16, "bold"), bg="#4CAF50", fg="white", padx=10, pady=5)
    select_button.pack(pady=20)

    export_button = tk.Button(frame, text="Export to Excel", command=export_to_excel, font=("Arial", 16, "bold"), bg="#FF5722", fg="white", padx=10, pady=5)
    export_button.pack(pady=20)

    # Run the Tkinter event loop
    root.mainloop()
    
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")
        raise e