# app.py
import streamlit as st
from utils.data_processor import DataProcessor
from utils.chart_updater import ChartUpdater
from utils.table_updater import TableUpdater
from utils.file_handler import FileHandler
from utils.historical_processor import HistoricalDataProcessor
import pandas as pd
import os
from datetime import datetime
import tempfile

# Turkish month names dictionary
TURKISH_MONTHS = {
    1: 'Oca', 2: 'Şub', 3: 'Mar', 4: 'Nis', 
    5: 'May', 6: 'Haz', 7: 'Tem', 8: 'Ağu',
    9: 'Eyl', 10: 'Eki', 11: 'Kas', 12: 'Ara'
}

def get_month_year_suffix():
    now = datetime.now()
    month = TURKISH_MONTHS[now.month]
    year = str(now.year)[2:]  # Get last two digits of year
    return f"{month}{year}"

def process_survey_data(survey_file, tr_output_path, en_output_path, historical_file_path, table_template_path):
    try:
        # Initialize processors
        file_handler = FileHandler()
        data_processor = DataProcessor()
        
        # Initialize chart updaters for both languages
        chart_updater_tr = ChartUpdater(tr_output_path, language='tr')
        chart_updater_en = ChartUpdater(en_output_path, language='en')
        
        historical_processor = HistoricalDataProcessor(historical_file_path)
        
        # Create output paths for tables with month-year suffix
        try:
            print("Setting up table updaters...")
            temp_dir = tempfile.gettempdir()
            now = datetime.now()
            month = TURKISH_MONTHS[now.month]
            year = str(now.year)[2:]
            month_year = f"{month}{year}"
            
            tr_table_output_path = os.path.join(temp_dir, f'Tables_{month_year}.xlsx')
            en_table_output_path = os.path.join(temp_dir, f'Tables_{month_year}_en.xlsx')
            print(f"Table outputs will be saved to: {tr_table_output_path} and {en_table_output_path}")
            
            # Initialize both Turkish and English table updaters
            tr_table_updater = TableUpdater(table_template_path, tr_table_output_path, language='tr')
            en_table_updater = TableUpdater(table_template_path, en_table_output_path, language='en')
            print("Successfully initialized table updaters")
        except Exception as e:
            raise Exception(f"Error setting up table updaters: {str(e)}")

        # Read and process survey data
        try:
            survey_df = pd.read_excel(survey_file)
            print(f"Successfully read survey data with {len(survey_df)} rows")
        except Exception as e:
            raise Exception(f"Error reading survey file: {str(e)}")
        
        # Set the parti column from the survey question
        try:
            survey_df['parti'] = survey_df["Bu Pazar genel seçim olsa hangi partiye oy verirsiniz?"]
            print("Successfully set parti column")
        except Exception as e:
            raise Exception(f"Error setting parti column: {str(e)}")
        
        processed_data = data_processor.process_survey_data(
            survey_df,
            "Bu Pazar genel seçim olsa hangi partiye oy verirsiniz?"
        )
        
        # Create education column
        def map_education(edu):
            if edu in ['Doktora', 'Yüksek lisans', 'Yüksekokul veya üniversite mezunu']:
                return 'Yüksekokul ve üzeri'
            elif edu == 'Lise ve dengi meslek okulu mezunu':
                return 'Lise'
            else:
                return 'İlköğretim ve altı'
        
        try:
            survey_df['education'] = survey_df['En son mezun olduğunuz eğitim kurumunu belirtir misiniz? Halihazırda eğitiminize devam ediyorsanız lütfen şu anda devam ettiğiniz eğitim seviyesini belirtin.'].apply(map_education)
            print("Successfully created education column")
        except Exception as e:
            raise Exception(f"Error creating education column: {str(e)}")
        
        # Create age group column
        def map_age_group(age):
            age = int(age)
            if 18 <= age <= 34:
                return '18-34'
            elif 35 <= age <= 54:
                return '35-54'
            else:
                return '55 ve üstü'
        
        try:
            # Find the age column
            age_col = historical_processor._find_column(survey_df, "Yaşınızı öğrenebilir miyim")
            survey_df['age_group_second'] = survey_df[age_col].apply(map_age_group)
            print("Successfully created age group column")
        except Exception as e:
            raise Exception(f"Error creating age group column: {str(e)}")
        
        # Process historical data
        try:
            historical_data = {}
            
            # Process party votes
            historical_data['party_votes'] = historical_processor.process_party_votes(survey_df)
            
            # Process education breakdown
            education_data = historical_processor.process_education_breakdown(survey_df)
            historical_data.update(education_data)
            
            # Process age breakdown
            age_data = historical_processor.process_age_breakdown(survey_df)
            historical_data.update(age_data)
            
            # Process 2023 party data
            historical_data['party_votes_2023'] = historical_processor.process_2023_party_breakdown(survey_df)
            
            # Process economic data
            historical_data['econ_main'] = historical_processor.process_econ_main(survey_df)
            historical_data['econ_negative_party'] = historical_processor.process_econ_negative_party(survey_df)
            historical_data['econ_negative_age'] = historical_processor.process_econ_negative_age(survey_df)
            historical_data['econ_negative_education'] = historical_processor.process_econ_negative_education(survey_df)
            historical_data['econ_future_main'] = historical_processor.process_econ_future_main(survey_df)
            historical_data['econ_future_party'] = historical_processor.process_econ_future_party(survey_df)
            historical_data['econ_future_age'] = historical_processor.process_econ_future_age(survey_df)
            
            # Process politician success data
            historical_data['current_success'] = historical_processor.process_politician_success(survey_df)
            historical_data['politician_success_main'] = historical_processor.process_politician_success_main(survey_df)
            historical_data['politician_success_second'] = historical_processor.process_politician_success_second(survey_df)
            
            # Process subsistence data
            historical_data['subsistence'] = historical_processor.process_subsistence(survey_df)
            historical_data['subsistence_party'] = historical_processor.process_subsistence_party(survey_df)
            
            print("Successfully processed historical data")
        except Exception as e:
            raise Exception(f"Error processing historical data: {str(e)}")
        
        # Save updated historical data
        try:
            for sheet_name, df in historical_data.items():
                historical_processor.save_updated_data(df, sheet_name)
            print("Successfully saved historical data")
        except Exception as e:
            raise Exception(f"Error saving historical data: {str(e)}")
        
        # Update all charts in both languages
        try:
            # Update Turkish version
            chart_updater_tr.update_all_charts(processed_data, historical_data)
            print("Successfully updated Turkish charts")
            
            # Update English version
            chart_updater_en.update_all_charts(processed_data, historical_data)
            print("Successfully updated English charts")
        except Exception as e:
            raise Exception(f"Error updating charts: {str(e)}")
        
        # Update tables in both languages
        try:
            print("Starting table updates...")
            print("Updating Turkish tables...")
            tr_table_updater.update_all_tables(survey_df)
            print("Successfully updated Turkish tables")
            
            print("Updating English tables...")
            en_table_updater.update_all_tables(survey_df)
            print("Successfully updated English tables")
        except Exception as e:
            raise Exception(f"Error updating tables: {str(e)}")
        
        # Return both Turkish and English file paths along with other results
        return True, "Data processed successfully", tr_table_output_path, en_table_output_path, historical_file_path, tr_output_path, en_output_path
        
    except Exception as e:
        return False, f"Error processing data: {str(e)}", None, None, None, None, None

def main():
    # Initialize session state for file paths if they don't exist
    if 'tr_output_path' not in st.session_state:
        st.session_state.tr_output_path = None
    if 'en_output_path' not in st.session_state:
        st.session_state.en_output_path = None
    if 'tr_table_output_path' not in st.session_state:
        st.session_state.tr_table_output_path = None
    if 'en_table_output_path' not in st.session_state:
        st.session_state.en_table_output_path = None
    if 'historical_output_path' not in st.session_state:
        st.session_state.historical_output_path = None

    # Initialize FileHandler
    file_handler = FileHandler()

    # Set page config
    st.set_page_config(
        page_title="Türkiye Raporu",
        page_icon="📊",
        layout="wide"
    )

    # Define path for table template
    table_template_path = "table_data/table_templates_main.xlsx"

    # Sidebar
    with st.sidebar:
        # Add logo
        st.image("https://www.turkiyeraporu.com/wp-content/uploads/2021/04/logo_yeni.png", width=300)
        st.markdown("---")
        
        st.markdown("### Upload Files")
        # File upload section - now for survey, template, and historical data
        template_file = st.file_uploader("Upload PowerPoint Template", type=['pptx'])
        survey_file = st.file_uploader("Upload Survey Data (Excel)", type=['xlsx'])
        historical_file = st.file_uploader("Upload Historical Data (Excel)", type=['xlsx'])

    # Main content
    #st.title("📊 Türkiye Raporu - Report Generator")
    
    # Description
    st.markdown("""
    #### How to Use:
    1. Upload your survey data Excel file (Current)
    2. Upload the PowerPoint template (Previous month's PowerPoint)
    3. Upload the historical data Excel file (Previous month's historical data)
    4. Click process to generate your report
    """)

    # Process button and results
    if all([survey_file, template_file, historical_file]):
        st.markdown("---")
        
        if st.button("🚀 Process Data", type="primary"):
            try:
                # Create new processed files
                template_path, tr_output_path, en_output_path = file_handler.create_processed_file(template_file)
                
                # Save historical file and get its path
                historical_path = file_handler.save_uploaded_file(historical_file)
                
                # Process the data
                with st.spinner('Processing data...'):
                    success, message, tr_table_output_path, en_table_output_path, historical_output_path, tr_output_path, en_output_path = process_survey_data(
                        survey_file, 
                        tr_output_path,  # Pass Turkish output path
                        en_output_path,  # Pass English output path
                        historical_path,
                        table_template_path
                    )
                
                if success:
                    st.success(message)
                    # Store paths in session state
                    st.session_state.tr_output_path = tr_output_path
                    st.session_state.en_output_path = en_output_path
                    st.session_state.tr_table_output_path = tr_table_output_path
                    st.session_state.en_table_output_path = en_table_output_path
                    st.session_state.historical_output_path = historical_output_path
                else:
                    st.error(message)
                    
            except Exception as e:
                st.error(f"Error processing files: {str(e)}")
        
        # Show download buttons if paths exist in session state
        if (st.session_state.tr_output_path or st.session_state.en_output_path or 
            st.session_state.tr_table_output_path or st.session_state.en_table_output_path or 
            st.session_state.historical_output_path):
            st.markdown("### Download Processed Files")
            
            # Create three columns for better organization
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.markdown("#### 📊 Presentation Files")
                if st.session_state.tr_output_path:
                    file_handler.get_download_button(st.session_state.tr_output_path, "📥 Turkish PowerPoint (TR)")
                if st.session_state.en_output_path:
                    file_handler.get_download_button(st.session_state.en_output_path, "📥 English PowerPoint (EN)")
            
            with col2:
                st.markdown("#### 📈 Excel Reports")
                if st.session_state.tr_table_output_path:
                    file_handler.get_download_button(st.session_state.tr_table_output_path, "📥 Turkish Tables")
                if st.session_state.en_table_output_path:
                    file_handler.get_download_button(st.session_state.en_table_output_path, "📥 English Tables")
            
            with col3:
                st.markdown("#### 📚 Historical Data")
                if st.session_state.historical_output_path:
                    file_handler.get_download_button(st.session_state.historical_output_path, "📚 Updated Historical Data")
    else:
        st.info("👆 Please upload the survey data, PowerPoint template, and historical data files to begin processing")

if __name__ == "__main__":
    main()