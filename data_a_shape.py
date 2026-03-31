import streamlit as st
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point
import tempfile
import os
import zipfile
from io import BytesIO
from typing import Optional, Dict, Any
import chardet
import openpyxl


class BiodiversityExcelReader:
    """Lector de Excel de complejidad alta"""
    
    def __init__(self):
        self.required_columns = ['gbifID', 'kingdom', 'phylum', 'class', 'order', 'family', 'genus', 'species']
        
    def read_file(self, file) -> Optional[Dict[str, Any]]:
        """
        Read Excel file and extract biodiversity data
        
        Args:
            file: Uploaded file object (StreamlitUploadedFile or similar)
            
        Returns:
            Dictionary with sheets data or None if error
        """
        try:
            # Read all sheets from Excel file
            excel_file = pd.ExcelFile(file)
            sheets_data = {}
            
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet_name, header=0)
                
                # Clean the dataframe
                df = self._clean_dataframe(df)
                
                # Store sheet info
                sheets_data[sheet_name] = {
                    'dataframe': df,
                    'row_count': len(df),
                    'column_count': len(df.columns),
                    'columns': list(df.columns),
                    'has_coordinates': self._has_coordinates(df),
                    'has_taxonomy': self._has_taxonomy_columns(df),
                    'stats': self._calculate_stats(df)
                }
            
            return {
                'success': True,
                'sheets': sheets_data,
                'file_name': file.name,
                'total_sheets': len(sheets_data),
                'total_records': sum(s['row_count'] for s in sheets_data.values())
            }
            
        except Exception as e:
            st.error(f"Error reading Excel file: {str(e)}")
            return None
    
    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean and prepare dataframe"""
        # Remove completely empty rows
        df = df.dropna(how='all')
        
        # Remove completely empty columns
        df = df.dropna(axis=1, how='all')
        
        # Reset index
        df = df.reset_index(drop=True)
        
        # Convert column names to strings and clean them
        df.columns = [str(col).strip() if pd.notna(col) else f'column_{i}' 
                     for i, col in enumerate(df.columns)]
        
        return df
    
    def _has_coordinates(self, df: pd.DataFrame) -> bool:
        """Check if dataframe has coordinate columns"""
        # Look for columns with latitude/longitude in name
        lat_cols = [col for col in df.columns if 'latitude' in col.lower() or 'lat' in col.lower()]
        lon_cols = [col for col in df.columns if 'longitude' in col.lower() or 'lon' in col.lower()]
    
        # If not found, look for columns that might contain coordinate data
        if len(lat_cols) == 0 or len(lon_cols) == 0:
            # Try to find columns with numeric values that look like coordinates
            for col in df.columns:
                # Sample some non-null values
                sample = df[col].dropna().head(10)
                if len(sample) > 0:
                    # Check if values look like coordinates (numbers with possible comma decimal)
                    sample_str = sample.astype(str)
                    if sample_str.str.contains(r'^-?\d+[,.]?\d*$').any():
                        # This column contains numbers that could be coordinates
                        # We'll need the user to specify later
                        pass
    
        return len(lat_cols) > 0 and len(lon_cols) > 0
    
    def _has_taxonomy_columns(self, df: pd.DataFrame) -> bool:
        """Check if dataframe has taxonomy columns"""
        df_cols_lower = [str(col).lower() for col in df.columns]
        return any(col in df_cols_lower for col in ['kingdom', 'phylum', 'class', 'order', 'family', 'genus', 'species'])
    
    def _calculate_stats(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Calculate statistics for the dataframe"""
        stats = {}
        
        # Count records by kingdom if available
        if 'kingdom' in df.columns:
            stats['kingdom_counts'] = df['kingdom'].value_counts().to_dict()
        
        # Count records by phylum if available
        if 'phylum' in df.columns:
            stats['phylum_counts'] = df['phylum'].value_counts().head(10).to_dict()
        
        # Count records by class if available
        if 'class' in df.columns:
            stats['class_counts'] = df['class'].value_counts().head(10).to_dict()
        
        # Count records by order if available
        if 'order' in df.columns:
            stats['order_counts'] = df['order'].value_counts().head(10).to_dict()
        
        # Count records by family if available
        if 'family' in df.columns:
            stats['family_counts'] = df['family'].value_counts().head(10).to_dict()
        
        # Count records by genus if available
        if 'genus' in df.columns:
            stats['genus_counts'] = df['genus'].value_counts().head(10).to_dict()
        
        # Count records by species if available
        if 'species' in df.columns:
            stats['species_counts'] = df['species'].value_counts().head(10).to_dict()
        
        # Check for coordinate completeness
        if self._has_coordinates(df):
            lat_col = next((col for col in df.columns if 'latitude' in col.lower() or 'lat' in col.lower()), None)
            lon_col = next((col for col in df.columns if 'longitude' in col.lower() or 'lon' in col.lower()), None)
            
            if lat_col and lon_col:
                stats['coordinates_complete'] = df[[lat_col, lon_col]].notna().all(axis=1).sum()
                stats['coordinates_missing'] = len(df) - stats['coordinates_complete']
        
        # Check for dates
        date_cols = [col for col in df.columns if 'date' in col.lower() or 'eventdate' in col.lower()]
        if date_cols:
            stats['has_dates'] = True
            stats['date_columns'] = date_cols
        
        return stats
    
    def get_preview(self, df: pd.DataFrame, n_rows: int = 10) -> pd.DataFrame:
        """Get preview of dataframe"""
        return df.head(n_rows)
    
    def filter_by_taxon(self, df: pd.DataFrame, taxon_level: str, taxon_name: str) -> pd.DataFrame:
        """Filter dataframe by taxon level"""
        if taxon_level in df.columns:
            return df[df[taxon_level].astype(str).str.contains(taxon_name, case=False, na=False)]
        return df
    
    def filter_by_coordinates(self, df: pd.DataFrame, lat_min: float, lat_max: float, 
                             lon_min: float, lon_max: float) -> pd.DataFrame:
        """Filter dataframe by coordinate bounds"""
        lat_col = next((col for col in df.columns if 'latitude' in col.lower() or 'lat' in col.lower()), None)
        lon_col = next((col for col in df.columns if 'longitude' in col.lower() or 'lon' in col.lower()), None)
        
        if lat_col and lon_col:
            return df[
                (df[lat_col].between(lat_min, lat_max)) &
                (df[lon_col].between(lon_min, lon_max))
            ]
        return df
    
    def export_to_csv(self, df: pd.DataFrame) -> bytes:
        """Export dataframe to CSV bytes"""
        output = BytesIO()
        df.to_csv(output, index=False, encoding='utf-8')
        return output.getvalue()
    
    def export_to_shapefile(self, df: pd.DataFrame, sheet_name: str) -> Optional[BytesIO]:
        """
        Export dataframe to Shapefile (as zip)
    
        Args:
            df: DataFrame to export
            sheet_name: Name of the sheet for file naming
        
        Returns:
            BytesIO object containing zip file with shapefile components, or None if error
        """
        try:
            # Check if coordinates exist
            if not self._has_coordinates(df):
                st.error("Cannot create shapefile: No coordinate columns found")
                return None
        
            # Find coordinate columns
            lat_col = next((col for col in df.columns if 'latitude' in col.lower() or 'lat' in col.lower()), None)
            lon_col = next((col for col in df.columns if 'longitude' in col.lower() or 'lon' in col.lower()), None)
        
            if not lat_col or not lon_col:
                st.error("Could not identify latitude and longitude columns")
                return None
        
            # Remove rows with missing coordinates
            df_clean = df.dropna(subset=[lat_col, lon_col]).copy()
        
            if len(df_clean) == 0:
                st.error("No records with valid coordinates found")
                return None
        
            # Convert coordinates to numeric (handle comma as decimal separator)
            # Replace comma with dot for decimal values
            df_clean[lat_col] = df_clean[lat_col].astype(str).str.replace(',', '.').astype(float)
            df_clean[lon_col] = df_clean[lon_col].astype(str).str.replace(',', '.').astype(float)
        
            # Remove any rows where coordinates couldn't be converted to numeric
            df_clean = df_clean.dropna(subset=[lat_col, lon_col])
        
            if len(df_clean) == 0:
                st.error("No records with valid numeric coordinates found")
                return None
        
            # Convert all datetime columns to string to avoid shapefile datetime issues
            for col in df_clean.columns:
                # Check if column contains datetime-like data
                if pd.api.types.is_datetime64_any_dtype(df_clean[col]):
                    df_clean[col] = df_clean[col].astype(str)
                elif df_clean[col].dtype == 'object':
                    # Check if any values in object column are datetime
                    try:
                        # Sample first few non-null values to check if they're datetime-like
                        sample = df_clean[col].dropna().iloc[:5]
                        if len(sample) > 0:
                            # Try to convert to datetime, if successful, convert to string
                            pd.to_datetime(sample, errors='raise')
                            df_clean[col] = df_clean[col].astype(str)
                    except (ValueError, TypeError, KeyError):
                        pass
        
            # Also convert any timedelta columns if they exist
            for col in df_clean.select_dtypes(include=['timedelta']).columns:
                df_clean[col] = df_clean[col].astype(str)
        
            # Convert any other problematic types
            for col in df_clean.select_dtypes(include=['complex', 'object']).columns:
                # Try to convert to string to avoid any issues
                df_clean[col] = df_clean[col].astype(str)
        
            # Create geometry column
            geometry = [Point(xy) for xy in zip(df_clean[lon_col], df_clean[lat_col])]
        
            # Create GeoDataFrame
            gdf = gpd.GeoDataFrame(df_clean, geometry=geometry, crs="EPSG:4326")
        
            # Verify that gdf is actually a GeoDataFrame
            if not isinstance(gdf, gpd.GeoDataFrame):
                st.error("Failed to create GeoDataFrame")
                return None
        
            # Clean column names for shapefile (max 10 chars, no special chars)
            # Also ensure no duplicate column names
            original_columns = gdf.columns.tolist()
            new_columns = []
            seen_names = set()
            
            for col in original_columns:
                # Skip geometry column for renaming
                if col == 'geometry':
                    new_columns.append('geometry')
                    continue
            
                # Clean the column name
                clean_col = self._clean_column_name(col)
            
                # Handle duplicates by adding a suffix
                counter = 1
                base_name = clean_col
                while clean_col in seen_names:
                    clean_col = f"{base_name}_{counter}"
                    counter += 1
            
                seen_names.add(clean_col)
                new_columns.append(clean_col)
        
            # Rename columns
            gdf.columns = new_columns
        
            # Create a temporary directory
            with tempfile.TemporaryDirectory() as tmp_dir:
                # Define shapefile path
                shp_name = f"{sheet_name[:8].replace(' ', '_')}_export"
                shp_path = os.path.join(tmp_dir, f"{shp_name}.shp")
            
                # Save shapefile with explicit encoding and handling
                try:
                    gdf.to_file(shp_path, driver='ESRI Shapefile', encoding='utf-8')
                except Exception as e:
                    st.error(f"Error saving shapefile: {str(e)}")
                    # Try again with different encoding
                    try:
                        gdf.to_file(shp_path, driver='ESRI Shapefile', encoding='latin1')
                    except Exception as e2:
                        st.error(f"Failed with latin1 encoding as well: {str(e2)}")
                        return None
            
                # Check if file was created
                if not os.path.exists(shp_path):
                    st.error("Shapefile was not created successfully")
                    return None
            
                # Create zip file in memory
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    # Add all shapefile components to zip
                    for file in os.listdir(tmp_dir):
                        file_path = os.path.join(tmp_dir, file)
                        zip_file.write(file_path, arcname=file)
            
                zip_buffer.seek(0)
                return zip_buffer
            
        except Exception as e:
            st.error(f"Error creating shapefile: {str(e)}")
            import traceback
            st.error(traceback.format_exc())
            return None
    
    def _clean_column_name(self, col_name: str) -> str:
        """Clean column name for shapefile compatibility"""
        # Convert to string
        col = str(col_name)
        
        # Remove special characters, keep only alphanumeric and underscore
        col = ''.join(c for c in col if c.isalnum() or c == '_')
        
        # Replace spaces with underscore
        col = col.replace(' ', '_')
        
        # Truncate to 10 characters (shapefile limit)
        col = col[:10]
        
        # Ensure it's not empty
        if not col or col[0].isdigit():
            col = 'col_' + col
        
        return col
    
    def get_taxonomic_tree(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Build taxonomic tree from dataframe"""
        tree = {}
        
        if all(col in df.columns for col in ['kingdom', 'phylum', 'class', 'order', 'family', 'genus']):
            for _, row in df.iterrows():
                kingdom = str(row.get('kingdom', 'Unknown'))
                phylum = str(row.get('phylum', 'Unknown'))
                class_ = str(row.get('class', 'Unknown'))
                order = str(row.get('order', 'Unknown'))
                family = str(row.get('family', 'Unknown'))
                genus = str(row.get('genus', 'Unknown'))
                species = str(row.get('species', 'Unknown'))
                
                # Build tree structure
                if kingdom not in tree:
                    tree[kingdom] = {}
                if phylum not in tree[kingdom]:
                    tree[kingdom][phylum] = {}
                if class_ not in tree[kingdom][phylum]:
                    tree[kingdom][phylum][class_] = {}
                if order not in tree[kingdom][phylum][class_]:
                    tree[kingdom][phylum][class_][order] = {}
                if family not in tree[kingdom][phylum][class_][order]:
                    tree[kingdom][phylum][class_][order][family] = {}
                if genus not in tree[kingdom][phylum][class_][order][family]:
                    tree[kingdom][phylum][class_][order][family][genus] = set()
                
                tree[kingdom][phylum][class_][order][family][genus].add(species)
        
        # Convert sets to lists for JSON serialization
        return self._sets_to_lists(tree)
    
    def _sets_to_lists(self, obj):
        """Recursively convert sets to lists"""
        if isinstance(obj, dict):
            return {k: self._sets_to_lists(v) for k, v in obj.items()}
        elif isinstance(obj, set):
            return list(obj)
        else:
            return obj
    
    def get_summary_stats(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Get summary statistics for display"""
        stats = {
            'total_records': len(df),
            'columns': len(df.columns),
            'column_names': list(df.columns),
            'data_types': df.dtypes.astype(str).to_dict(),
            'missing_values': df.isna().sum().to_dict(),
            'unique_counts': {}
        }
        
        # Get unique counts for key columns
        for col in ['kingdom', 'phylum', 'class', 'order', 'family', 'genus', 'species']:
            if col in df.columns:
                stats['unique_counts'][col] = df[col].nunique()
        
        return stats

# Streamlit UI Integration
def render_excel_uploader():
    """Render Excel file uploader in Streamlit"""
    st.subheader("📊 Upload Excel File")
    
    reader = BiodiversityExcelReader()
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload your Excel file containing biodiversity data"
    )
    
    if uploaded_file is not None:
        with st.spinner("Reading file..."):
            result = reader.read_file(uploaded_file)
            
        if result and result['success']:
            st.success(f"✅ Successfully read {result['file_name']}")
            st.info(f"📈 Sheets: {result['total_sheets']} | Total records: {result['total_records']}")
            
            # Sheet selector
            selected_sheet = st.selectbox(
                "Select sheet to view",
                options=list(result['sheets'].keys())
            )
            
            sheet_data = result['sheets'][selected_sheet]
            df = sheet_data['dataframe']
            
            # Display sheet info
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Rows", sheet_data['row_count'])
            with col2:
                st.metric("Columns", sheet_data['column_count'])
            with col3:
                st.metric("Has Coordinates", "✅ Yes" if sheet_data['has_coordinates'] else "❌ No")
            
            # Tabs for different views
            tab1, tab2, tab3, tab4 = st.tabs(["📋 Preview Data", "📊 Statistics", "🌳 Taxonomy", "🔍 Filter & Export"])
            
            with tab1:
                st.dataframe(reader.get_preview(df), use_container_width=True)
            
            with tab2:
                stats = reader.get_summary_stats(df)
                
                # Summary metrics
                st.subheader("Summary")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Records", stats['total_records'])
                with col2:
                    st.metric("Total Columns", stats['columns'])
                with col3:
                    st.metric("Complete Records", 
                             df.dropna().shape[0] if len(df) > 0 else 0)
                
                # Unique counts
                if stats['unique_counts']:
                    st.subheader("Unique Values")
                    for level, count in stats['unique_counts'].items():
                        st.write(f"**{level.capitalize()}:** {count} unique")
                
                # Column info
                st.subheader("Column Information")
                col_info = pd.DataFrame({
                    'Column': stats['column_names'],
                    'Data Type': [stats['data_types'][col] for col in stats['column_names']],
                    'Missing': [stats['missing_values'][col] for col in stats['column_names']],
                    'Missing %': [round(stats['missing_values'][col]/len(df)*100, 1) 
                                 if len(df) > 0 else 0 for col in stats['column_names']]
                })
                st.dataframe(col_info, use_container_width=True)
            
            with tab3:
                tree = reader.get_taxonomic_tree(df)
                if tree:
                    st.json(tree)
                else:
                    st.warning("Insufficient taxonomic data to build tree")
            
            with tab4:
                st.subheader("Filter Data")
                
                # Initialize filtered_df with the original dataframe
                filtered_df = df.copy()
                filter_applied = False
                
                # Taxon filter
                if sheet_data['has_taxonomy']:
                    st.markdown("**Filter by Taxonomy**")
                    
                    # Create columns for filter options
                    filter_col1, filter_col2 = st.columns([1, 2])
                    
                    with filter_col1:
                        taxon_level = st.selectbox(
                            "Taxonomic Level",
                            options=['kingdom', 'phylum', 'class', 'order', 'family', 'genus', 'species'],
                            format_func=lambda x: x.capitalize(),
                            key="taxon_filter"
                        )
                    
                    if taxon_level in df.columns:
                        # Get all unique values, sorted
                        unique_values = sorted(df[taxon_level].dropna().unique())
                        total_unique = len(unique_values)
                        
                        with filter_col2:
                            # Add search functionality for species level
                            if taxon_level == 'species' and total_unique > 50:
                                st.markdown(f"**Total unique {taxon_level}s: {total_unique}**")
                                
                                # Search box for species
                                search_term = st.text_input(
                                    f"Search {taxon_level}",
                                    placeholder=f"Type to search among {total_unique} species...",
                                    key="species_search"
                                )
                                
                                if search_term:
                                    # Filter unique values by search term
                                    filtered_values = [val for val in unique_values 
                                                      if search_term.lower() in str(val).lower()]
                                    display_values = filtered_values
                                    st.info(f"Found {len(display_values)} matching species")
                                else:
                                    # Show a more manageable list (first 100 with option to show more)
                                    show_all = st.checkbox("Show all species", key="show_all_species")
                                    if show_all:
                                        display_values = unique_values
                                        st.info(f"Showing all {len(display_values)} species")
                                    else:
                                        display_values = unique_values[:100]
                                        st.info(f"Showing first 100 of {total_unique} species. Check 'Show all species' to see more, or use search to find specific species.")
                            else:
                                # For other taxonomic levels or smaller datasets
                                display_values = unique_values
                                if total_unique > 50:
                                    st.info(f"Showing all {total_unique} {taxon_level}s")
                            
                            if display_values:
                                selected_taxon = st.selectbox(
                                    f"Select {taxon_level}",
                                    options=display_values,
                                    key="taxon_value"
                                )
                                
                                if selected_taxon:
                                    filtered_df = reader.filter_by_taxon(filtered_df, taxon_level, selected_taxon)
                                    filter_applied = True
                                    st.success(f"✅ Applied {taxon_level} filter: {len(filtered_df)} records found")
                            else:
                                st.warning(f"No {taxon_level}s found matching your search")
                
                # Coordinate filter
                if sheet_data['has_coordinates']:
                    st.markdown("**Spatial Filter**")
                    
                    lat_col = next((col for col in df.columns if 'latitude' in col.lower() or 'lat' in col.lower()), None)
                    lon_col = next((col for col in df.columns if 'longitude' in col.lower() or 'lon' in col.lower()), None)
                    
                    if lat_col and lon_col:
                        col1, col2 = st.columns(2)
                        with col1:
                            lat_min = st.number_input("Min Latitude", value=-90.0, step=1.0)
                            lat_max = st.number_input("Max Latitude", value=90.0, step=1.0)
                        with col2:
                            lon_min = st.number_input("Min Longitude", value=-180.0, step=1.0)
                            lon_max = st.number_input("Max Longitude", value=180.0, step=1.0)
                        
                        if st.button("Apply Spatial Filter", key="apply_spatial"):
                            filtered_df = reader.filter_by_coordinates(filtered_df, lat_min, lat_max, lon_min, lon_max)
                            filter_applied = True
                            st.success(f"✅ Applied spatial filter: {len(filtered_df)} records found")
                
                # Show current filter status
                st.markdown("---")
                if filter_applied:
                    st.success(f"**Current filtered records: {len(filtered_df)}**")
                else:
                    st.info(f"**No filters applied. Total records: {len(filtered_df)}**")
                
                # Show filtered data preview
                st.subheader("Filtered Data Preview")
                if len(filtered_df) > 0:
                    st.dataframe(reader.get_preview(filtered_df), use_container_width=True)
                else:
                    st.warning("No records match the current filters")
                
                # Export section at the bottom of filter tab
                st.markdown("---")
                st.subheader("💾 Export Filtered Data")
                
                if len(filtered_df) == 0:
                    st.warning("No data to export. Please adjust filters to include at least one record.")
                else:
                    # Export format selection
                    export_format = st.radio(
                        "Select export format",
                        options=["CSV", "Shapefile (ZIP)"],
                        horizontal=True,
                        key="export_format"
                    )
                    
                    if export_format == "CSV":
                        st.markdown("**Export as CSV**")
                        csv_data = reader.export_to_csv(filtered_df)
                        st.download_button(
                            label="📥 Download CSV",
                            data=csv_data,
                            file_name=f"{selected_sheet}_filtered_export.csv",
                            mime="text/csv",
                            key="csv_download"
                        )
                        st.info("CSV file contains all filtered records including those without coordinates")
                    
                    else:  # Shapefile
                        st.markdown("**Export as Shapefile**")
                        
                        if not sheet_data['has_coordinates']:
                            st.warning("⚠️ Cannot create shapefile: No coordinate columns found in this sheet")
                        else:
                            # Show coordinate columns being used
                            lat_col = next((col for col in df.columns if 'latitude' in col.lower() or 'lat' in col.lower()), None)
                            lon_col = next((col for col in df.columns if 'longitude' in col.lower() or 'lon' in col.lower()), None)
                            
                            st.info(f"Using columns for coordinates:\n"
                                   f"- Latitude: **{lat_col}**\n"
                                   f"- Longitude: **{lon_col}**")
                            
                            # Count valid records in filtered data
                            valid_coords = filtered_df[[lat_col, lon_col]].notna().all(axis=1).sum()
                            st.write(f"Records with valid coordinates in filtered data: **{valid_coords}** / {len(filtered_df)}")
                            
                            if valid_coords == 0:
                                st.error("No records with valid coordinates to export. Please ensure your filtered data contains records with both latitude and longitude values.")
                            else:
                                if st.button("🔄 Generate Shapefile", key="generate_shp"):
                                    with st.spinner("Creating shapefile..."):
                                        zip_buffer = reader.export_to_shapefile(filtered_df, selected_sheet)
                                        
                                        if zip_buffer:
                                            st.download_button(
                                                label="📥 Download Shapefile (ZIP)",
                                                data=zip_buffer,
                                                file_name=f"{selected_sheet}_filtered_shapefile.zip",
                                                mime="application/zip",
                                                key="shp_download",
                                                disabled=False
                                            )
                                            st.success("✅ Shapefile created successfully!")
                                            st.info("**Note:** Shapefiles have limitations:\n"
                                                   "- Column names truncated to 10 characters\n"
                                                   "- Special characters removed from column names\n"
                                                   "- Only records with valid coordinates are included\n"
                                                   "- Download contains all shapefile components (.shp, .shx, .dbf, etc.)")







# Initialize session state
if 'excel_reader' not in st.session_state:
    st.session_state.excel_reader = BiodiversityExcelReader()

st.set_page_config(
    page_title="Excel to Shapefile Converter",
    page_icon="🗺️",
    layout="wide"
)

st.title("🗺️ Excel to Shapefile Converter")
st.markdown("Upload an Excel file containing biodiversity data with coordinates to convert to Shapefile format")

# Main app
render_excel_uploader()