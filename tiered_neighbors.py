import pandas as pd
import numpy as np
from math import radians, sin, cos, sqrt, atan2

# Haversine distance calculation (km)
def haversine(lat1, lon1, lat2, lon2):
    R = 6371  # Earth radius in km
    dlat, dlon = radians(lat2-lat1), radians(lon2-lon1)
    a = sin(dlat/2)**2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon/2)**2
    c = 2 * atan2(sqrt(a), sqrt(1-a))
    return R * c

print("Loading Excel file...")
# Load the Excel file
df = pd.read_excel('Sector-Addition-NBR-List.xlsx')

print("Separating source and target cells...")
# Separate source cells (rows with Source Cell ID) and target cells (rows with Sec Name)
sources = df[df['Source Cell ID'].notna()][['Source Cell ID', 'AZM', 'Lat', 'Long']].dropna()
sources.columns = ['Source_Cell', 'Src_Azm', 'Src_Lat', 'Src_Lon']

targets = df[df['Sec Name'].notna()][['Sec Name', 'Lat.1', 'Long.1', '4G Azm']].dropna()
targets.columns = ['Target_Cell', 'Tgt_Lat', 'Tgt_Lon', 'Tgt_Azm']

print(f"Found {len(sources)} source cells and {len(targets)} target cells")

# Generate neighbor list
nbr_list = []
print("Calculating neighbors...")

for idx_src, src in sources.iterrows():
    for idx_tgt, tgt in targets.iterrows():
        # Calculate distance and angle difference
        dist = haversine(src['Src_Lat'], src['Src_Lon'], tgt['Tgt_Lat'], tgt['Tgt_Lon'])
        
        # Calculate minimum angle difference (100° beam width: ±50°)
        angle_diff = abs((tgt['Tgt_Azm'] - src['Src_Azm'] + 180) % 360 - 180)
        
        # Filter: within 50° angle AND not self (dist > 0.1 km to avoid same site)
        if angle_diff <= 50 and dist > 0.1:
            # Area type determination (urban: 3km, rural: 5-10km)
            # Using distance-based proxy (customize with GIS data if available)
            area_type = 'urban' if dist < 4 else 'rural'
            max_dist = 3 if area_type == 'urban' else 10
            
            # Only include if within distance limit
            if dist <= max_dist:
                # Tier assignment based on distance
                if dist <= 3:
                    tier = '1st tier'
                elif dist <= 6:
                    tier = '2nd tier'
                elif dist <= 10:
                    tier = '3rd tier'
                else:
                    tier = 'Beyond'  # Shouldn't reach here due to max_dist filter
                
                nbr_list.append({
                    'Source_Cell': src['Source_Cell'],
                    'Target_Cell': tgt['Target_Cell'],
                    'Distance_km': round(dist, 2),
                    'Src_Azm': round(src['Src_Azm'], 1),
                    'Tgt_Azm': round(tgt['Tgt_Azm'], 1),
                    'Angle_Diff': round(angle_diff, 1),
                    'Area_Type': area_type,
                    'Tier': tier,
                    'Src_Lat': round(src['Src_Lat'], 6),
                    'Src_Lon': round(src['Src_Lon'], 6),
                    'Tgt_Lat': round(tgt['Tgt_Lat'], 6),
                    'Tgt_Lon': round(tgt['Tgt_Lon'], 6)
                })

# Create final DataFrame
nbr_df = pd.DataFrame(nbr_list)

if not nbr_df.empty:
    # Sort by Source_Cell and Distance (closest neighbors first)
    nbr_df = nbr_df.sort_values(['Source_Cell', 'Distance_km'])
    
    # Save to Excel with formatting
    with pd.ExcelWriter('Generated_NBR_List.xlsx', engine='openpyxl') as writer:
        nbr_df.to_excel(writer, sheet_name='Neighbor_List', index=False)
        
        # Auto-adjust column widths
        worksheet = writer.sheets['Neighbor_List']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 20)
            worksheet.column_dimensions[column_letter].width = adjusted_width

    print("\n" + "="*60)
    print("✅ NEIGHBOR LIST GENERATED SUCCESSFULLY!")
    print("="*60)
    print(f"📊 Total neighbor relations: {len(nbr_df):,}")
    print(f"🏘️  Urban neighbors: {len(nbr_df[nbr_df['Area_Type']=='urban']):,}")
    print(f"🌾 Rural neighbors: {len(nbr_df[nbr_df['Area_Type']=='rural']):,}")
    
    print("\n🏆 Tier Distribution:")
    print(nbr_df['Tier'].value_counts().sort_index())
    
    print("\n📋 Sample Output (Top 20 neighbors):")
    print("-" * 100)
    print(nbr_df.head(20)[['Source_Cell', 'Target_Cell', 'Distance_km', 'Tier', 'Area_Type']].to_string(index=False))
    
    print("\n📈 Top 10 Source Cells by Neighbor Count:")
    print(nbr_df['Source_Cell'].value_counts().head(10))
    
else:
    print("❌ No neighbors found matching criteria!")
    print("Check your distance thresholds or azimuth ranges.")

print("\n💾 Output saved to: Generated_NBR_List.xlsx")
print("\n🎯 Ready for OSS/BSS neighbor addition!")
