# utils/historical_processor.py
import pandas as pd
from datetime import datetime
from utils.date_formatter import TurkishDateFormatter
import numpy as np

class HistoricalDataProcessor:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.date_formatter = TurkishDateFormatter()
        self.sheet_names = {
            'party_votes': 'party_votes',
            'education_akp': 'party_votes_education_akp',
            'education_chp': 'party_votes_education_chp',
            'education_dem': 'party_votes_education_dem',
            'education_iyip': 'party_votes_education_iyip',
            'education_mhp': 'party_votes_education_mhp',
            'education_kararsiz': 'party_votes_education_kararsiz',
            'education_absent': 'party_votes_education_absent',
            'age_akp': 'party_votes_age_akp',
            'age_chp': 'party_votes_age_chp',
            'age_dem': 'party_votes_age_dem',
            'age_iyip': 'party_votes_age_iyip',
            'age_mhp': 'party_votes_age_mhp',
            'age_kararsiz': 'party_votes_age_kararsiz',
            'age_absent': 'party_votes_age_absent',
            '2023': 'party_votes_2023',
            # New economic sheets
            'econ_main': 'econ_main',
            'econ_negative_party': 'econ_negative_party',
            'econ_negative_age': 'econ_negative_age',
            'econ_negative_education': 'econ_negative_education',
            'econ_future_main': 'econ_future_main',
            'econ_future_party': 'econ_future_party',
            'econ_future_age': 'econ_future_age',
            # Politician success sheets
            'politician_success_main': 'politician_success_main',
            'politician_success_second': 'politician_success_second',
            # Subsistence sheets
            'subsistence': 'subsistence',
            'subsistence_party': 'subsistence_party'
        }
        
        self.party_mapping = {
            'Adalet ve Kalkınma Partisi (AKP)': 'AK Parti',
            'Cumhuriyet Halk Partisi (CHP)': 'CHP',
            'Yeşil Sol Parti (YSP)/ Halkların Demokratik Partisi (HDP)': 'DEM Parti',
            'İYİ Parti': 'İYİ Parti',
            'Milliyetçi Hareket Partisi (MHP)': 'MHP',
            'Kararsızım': 'Kararsız',
            'Oy kullanmayacağım': 'Oy Kullanmam'
        }
        
        self.party_mapping_2023 = {
            'Adalet ve Kalkınma Partisi (AK Parti/AKP)': 'AK Parti',
            'Cumhuriyet Halk Partisi (CHP)': 'CHP',
            'Yeşil Sol Parti (YSP) / Halkların Demokratik Partisi (HDP) / DEM Parti': 'DEM Parti',
            'İYİ Parti': 'İYİ Parti',
            'Milliyetçi Hareket Partisi (MHP)': 'MHP'
        }
        
        # New mappings for economic responses
        self.econ_current_mapping = {
            'Çok iyi': 'Çok İyi / İyi',
            'İyi': 'Çok İyi / İyi',
            'Ne iyi ne kötü': 'Ne iyi ne kötü',
            'Kötü': 'Çok kötü / Kötü',
            'Çok kötü': 'Çok kötü / Kötü'
        }
        
        self.econ_future_mapping = {
            'Çok daha iyi': 'Çok Daha İyi/Daha İyi',
            'Daha iyi': 'Çok Daha İyi/Daha İyi',
            'Değişmez': 'Değişmez',
            'Daha kötü': 'Çok Daha Kötü/Daha Kötü',
            'Çok daha kötü': 'Çok Daha Kötü/Daha Kötü'
        }
        
    def read_historical_data(self, sheet_name: str = 'party_votes') -> pd.DataFrame:
        """Read historical data from the specified sheet"""
        try:
            df = pd.read_excel(self.file_path, sheet_name=sheet_name)
            return df
        except Exception as e:
            print(f"Error reading historical data from sheet {sheet_name}: {str(e)}")
            return pd.DataFrame()

    def process_party_votes(self, survey_data: pd.DataFrame) -> pd.DataFrame:
        """Process main party votes"""
        df = self.read_historical_data('party_votes')
        current_date = self.date_formatter.format_date(datetime.now())
        
        # Map party names
        survey_data['mapped_party'] = survey_data['parti'].map(self.party_mapping)
        
        # Calculate percentages
        total_weight = survey_data['duzeltilmis_agirlik'].sum()
        party_percentages = {}
        
        for party in ['AK Parti', 'CHP', 'DEM Parti', 'İYİ Parti', 'MHP', 'Kararsız', 'Oy Kullanmam']:
            party_data = survey_data[survey_data['mapped_party'] == party]
            percentage = (party_data['duzeltilmis_agirlik'].sum() / total_weight * 100) if total_weight > 0 else 0
            party_percentages[party] = percentage
        
        # Calculate Diğer
        other_parties = survey_data[~survey_data['mapped_party'].isin(party_percentages.keys())]
        party_percentages['Diğer'] = (other_parties['duzeltilmis_agirlik'].sum() / total_weight * 100) if total_weight > 0 else 0
        
        # Create new row
        if df.empty:
            df = pd.DataFrame(columns=['Months'] + list(party_percentages.keys()))
        
        df.loc[len(df)] = [current_date] + [party_percentages[col] for col in df.columns[1:]]
        return df

    def process_education_breakdown(self, survey_data: pd.DataFrame) -> dict:
        """Process education breakdown for each party"""
        current_date = self.date_formatter.format_date(datetime.now())
        results = {}
        
        # Map party names
        survey_data['mapped_party'] = survey_data['parti'].map(self.party_mapping)
        
        # First calculate totals for each education level
        education_totals = {}
        for education_level in ['İlköğretim ve altı', 'Lise', 'Yüksekokul ve üzeri']:
            education_totals[education_level] = survey_data[survey_data['education'] == education_level]['duzeltilmis_agirlik'].sum()
        
        for party_original, party_mapped in self.party_mapping.items():
            sheet_suffix = {
                'AK Parti': 'akp',
                'CHP': 'chp',
                'DEM Parti': 'dem',
                'İYİ Parti': 'iyip',
                'MHP': 'mhp',
                'Kararsız': 'kararsiz',
                'Oy Kullanmam': 'absent'
            }.get(party_mapped)
            
            if not sheet_suffix:
                continue
                
            sheet_name = f'party_votes_education_{sheet_suffix}'
            df = self.read_historical_data(sheet_name)
            
            # Calculate percentages by education level (column percentages)
            party_data = survey_data[survey_data['mapped_party'] == party_mapped]
            education_percentages = {}
            
            for education_level in ['İlköğretim ve altı', 'Lise', 'Yüksekokul ve üzeri']:
                level_total = education_totals[education_level]
                level_data = party_data[party_data['education'] == education_level]
                percentage = (level_data['duzeltilmis_agirlik'].sum() / level_total * 100) if level_total > 0 else 0
                education_percentages[education_level] = percentage
            
            # Create or update DataFrame
            if df.empty:
                df = pd.DataFrame(columns=['Months'] + list(education_percentages.keys()))
            
            df.loc[len(df)] = [current_date] + [education_percentages[col] for col in df.columns[1:]]
            results[sheet_name] = df
            
        return results

    def process_age_breakdown(self, survey_data: pd.DataFrame) -> dict:
        """Process age breakdown for each party"""
        current_date = self.date_formatter.format_date(datetime.now())
        results = {}
        
        # Map party names
        survey_data['mapped_party'] = survey_data['parti'].map(self.party_mapping)
        
        # First calculate totals for each age group
        age_totals = {}
        for age_group in ['18-34', '35-54', '55 ve üstü']:
            age_totals[age_group] = survey_data[survey_data['age_group_second'] == age_group]['duzeltilmis_agirlik'].sum()
        
        for party_original, party_mapped in self.party_mapping.items():
            sheet_suffix = {
                'AK Parti': 'akp',
                'CHP': 'chp',
                'DEM Parti': 'dem',
                'İYİ Parti': 'iyip',
                'MHP': 'mhp',
                'Kararsız': 'kararsiz',
                'Oy Kullanmam': 'absent'
            }.get(party_mapped)
            
            if not sheet_suffix:
                continue
                
            sheet_name = f'party_votes_age_{sheet_suffix}'
            df = self.read_historical_data(sheet_name)
            
            # Calculate percentages by age group (column percentages)
            party_data = survey_data[survey_data['mapped_party'] == party_mapped]
            age_percentages = {}
            
            for age_group in ['18-34', '35-54', '55 ve üstü']:
                group_total = age_totals[age_group]
                group_data = party_data[party_data['age_group_second'] == age_group]
                percentage = (group_data['duzeltilmis_agirlik'].sum() / group_total * 100) if group_total > 0 else 0
                age_percentages[age_group] = percentage
            
            # Create or update DataFrame
            if df.empty:
                df = pd.DataFrame(columns=['Months'] + list(age_percentages.keys()))
            
            df.loc[len(df)] = [current_date] + [age_percentages[col] for col in df.columns[1:]]
            results[sheet_name] = df
            
        return results

    def process_2023_party_breakdown(self, survey_data: pd.DataFrame) -> pd.DataFrame:
        """Process 2023 party breakdown data"""
        df = self.read_historical_data('party_votes_2023')
        current_date = self.date_formatter.format_date(datetime.now())
        
        # Map current and 2023 party names
        survey_data['mapped_party'] = survey_data['parti'].map(self.party_mapping)
        survey_data['mapped_2023_party'] = survey_data['2023 Genel Seçimlerinde hangi partiye oy verdiniz?'].map(self.party_mapping_2023)
        
        # Calculate percentages for each 2023 party choice
        party_percentages = {}
        for party_2023 in ['AK Parti', 'CHP', 'DEM Parti', 'İYİ Parti', 'MHP']:
            party_2023_data = survey_data[survey_data['mapped_2023_party'] == party_2023]
            total_weight = party_2023_data['duzeltilmis_agirlik'].sum()
            
            if total_weight > 0:
                current_party_data = party_2023_data[party_2023_data['mapped_party'] == party_2023]
                percentage = (current_party_data['duzeltilmis_agirlik'].sum() / total_weight * 100)
            else:
                percentage = 0
                
            party_percentages[party_2023] = percentage
        
        # Create or update DataFrame
        if df.empty:
            df = pd.DataFrame(columns=['Months'] + list(party_percentages.keys()))
        
        df.loc[len(df)] = [current_date] + [party_percentages[col] for col in df.columns[1:]]
        return df

    def _find_column(self, df: pd.DataFrame, search_text: str) -> str:
        """Find the exact column name that contains the given text"""
        matching_cols = [col for col in df.columns if search_text in col]
        if matching_cols:
            return matching_cols[0]
        raise ValueError(f"Could not find column containing: {search_text}")

    def process_econ_main(self, survey_data: pd.DataFrame) -> pd.DataFrame:
        """Process main economic situation data"""
        df = self.read_historical_data('econ_main')
        current_date = self.date_formatter.format_date(datetime.now())
        
        # Find the correct column name
        econ_current_col = self._find_column(survey_data, "Bugün itibari ile ekonominin nasıl olduğunu düşünüyorsunuz")
        
        # Map economic responses
        survey_data['econ_current_mapped'] = survey_data[econ_current_col].map(self.econ_current_mapping)
        
        # Calculate total weight
        total_weight = survey_data['duzeltilmis_agirlik'].sum()
        
        # Calculate percentages for each response group (row percentages)
        percentages = {}
        for response in ['Çok kötü / Kötü', 'Ne iyi ne kötü', 'Çok İyi / İyi']:
            response_data = survey_data[survey_data['econ_current_mapped'] == response]
            percentage = (response_data['duzeltilmis_agirlik'].sum() / total_weight * 100) if total_weight > 0 else 0
            percentages[response] = percentage
        
        # Create or update DataFrame
        if df.empty:
            df = pd.DataFrame(columns=['Months'] + list(percentages.keys()))
        
        df.loc[len(df)] = [current_date] + [percentages[col] for col in df.columns[1:]]
        return df

    def process_econ_negative_party(self, survey_data: pd.DataFrame) -> pd.DataFrame:
        """Process economic situation breakdown by party (negative responses only)"""
        df = self.read_historical_data('econ_negative_party')
        current_date = self.date_formatter.format_date(datetime.now())
        
        # Find the correct column name
        econ_current_col = self._find_column(survey_data, "Bugün itibari ile ekonominin nasıl olduğunu düşünüyorsunuz")
        
        # Map party names and economic responses
        survey_data['mapped_party'] = survey_data['2023 Genel Seçimlerinde hangi partiye oy verdiniz?'].map(self.party_mapping_2023)
        survey_data['econ_current_mapped'] = survey_data[econ_current_col].map(self.econ_current_mapping)
        
        # Filter for negative responses only
        negative_responses = survey_data[survey_data['econ_current_mapped'] == 'Çok kötü / Kötü']
        
        # Calculate percentages for each party
        party_percentages = {}
        for party in ['AK Parti', 'CHP', 'DEM Parti', 'İYİ Parti', 'MHP']:
            party_total = survey_data[survey_data['mapped_party'] == party]['duzeltilmis_agirlik'].sum()
            party_negative = negative_responses[negative_responses['mapped_party'] == party]['duzeltilmis_agirlik'].sum()
            percentage = (party_negative / party_total * 100) if party_total > 0 else 0
            party_percentages[party] = percentage
        
        # Create or update DataFrame
        if df.empty:
            df = pd.DataFrame(columns=['Months'] + list(party_percentages.keys()))
        
        df.loc[len(df)] = [current_date] + [party_percentages[col] for col in df.columns[1:]]
        return df

    def process_econ_negative_age(self, survey_data: pd.DataFrame) -> pd.DataFrame:
        """Process economic situation breakdown by age (negative responses only)"""
        df = self.read_historical_data('econ_negative_age')
        current_date = self.date_formatter.format_date(datetime.now())
        
        # Find the correct column name
        econ_current_col = self._find_column(survey_data, "Bugün itibari ile ekonominin nasıl olduğunu düşünüyorsunuz")
        
        # Map economic responses
        survey_data['econ_current_mapped'] = survey_data[econ_current_col].map(self.econ_current_mapping)
        
        # Filter for negative responses only
        negative_responses = survey_data[survey_data['econ_current_mapped'] == 'Çok kötü / Kötü']
        
        # Calculate percentages for each age group
        age_percentages = {}
        for age_group in ['18-34', '35-54', '55 ve üstü']:
            age_total = survey_data[survey_data['age_group_second'] == age_group]['duzeltilmis_agirlik'].sum()
            age_negative = negative_responses[negative_responses['age_group_second'] == age_group]['duzeltilmis_agirlik'].sum()
            percentage = (age_negative / age_total * 100) if age_total > 0 else 0
            age_percentages[age_group] = percentage
        
        # Create or update DataFrame
        if df.empty:
            df = pd.DataFrame(columns=['Months'] + list(age_percentages.keys()))
        
        df.loc[len(df)] = [current_date] + [age_percentages[col] for col in df.columns[1:]]
        return df

    def process_econ_negative_education(self, survey_data: pd.DataFrame) -> pd.DataFrame:
        """Process economic situation breakdown by education (negative responses only)"""
        df = self.read_historical_data('econ_negative_education')
        current_date = self.date_formatter.format_date(datetime.now())
        
        # Find the correct column name
        econ_current_col = self._find_column(survey_data, "Bugün itibari ile ekonominin nasıl olduğunu düşünüyorsunuz")
        
        # Map economic responses
        survey_data['econ_current_mapped'] = survey_data[econ_current_col].map(self.econ_current_mapping)
        
        # Filter for negative responses only
        negative_responses = survey_data[survey_data['econ_current_mapped'] == 'Çok kötü / Kötü']
        
        # Calculate percentages for each education level
        education_percentages = {}
        for education_level in ['İlköğretim ve altı', 'Lise', 'Yüksekokul ve üzeri']:
            edu_total = survey_data[survey_data['education'] == education_level]['duzeltilmis_agirlik'].sum()
            edu_negative = negative_responses[negative_responses['education'] == education_level]['duzeltilmis_agirlik'].sum()
            percentage = (edu_negative / edu_total * 100) if edu_total > 0 else 0
            education_percentages[education_level] = percentage
        
        # Create or update DataFrame
        if df.empty:
            df = pd.DataFrame(columns=['Months'] + list(education_percentages.keys()))
        
        df.loc[len(df)] = [current_date] + [education_percentages[col] for col in df.columns[1:]]
        return df

    def process_econ_future_main(self, survey_data: pd.DataFrame) -> pd.DataFrame:
        """Process main future economic situation data"""
        df = self.read_historical_data('econ_future_main')
        current_date = self.date_formatter.format_date(datetime.now())
        
        # Find the correct column name
        econ_future_col = self._find_column(survey_data, "Önümüzdeki bir yıl içerisinde ekonominin nasıl olacağını düşünüyorsunuz")
        
        # Map economic responses
        survey_data['econ_future_mapped'] = survey_data[econ_future_col].map(self.econ_future_mapping)
        
        # Calculate total weight
        total_weight = survey_data['duzeltilmis_agirlik'].sum()
        
        # Calculate percentages for each response group (row percentages)
        percentages = {}
        for response in ['Çok Daha Kötü/Daha Kötü', 'Değişmez', 'Çok Daha İyi/Daha İyi']:
            response_data = survey_data[survey_data['econ_future_mapped'] == response]
            percentage = (response_data['duzeltilmis_agirlik'].sum() / total_weight * 100) if total_weight > 0 else 0
            percentages[response] = percentage
        
        # Create or update DataFrame
        if df.empty:
            df = pd.DataFrame(columns=['Months'] + list(percentages.keys()))
        
        df.loc[len(df)] = [current_date] + [percentages[col] for col in df.columns[1:]]
        return df

    def process_econ_future_party(self, survey_data: pd.DataFrame) -> pd.DataFrame:
        """Process future economic situation breakdown by party (negative responses only)"""
        df = self.read_historical_data('econ_future_party')
        current_date = self.date_formatter.format_date(datetime.now())
        
        # Find the correct column name
        econ_future_col = self._find_column(survey_data, "Önümüzdeki bir yıl içerisinde ekonominin nasıl olacağını düşünüyorsunuz")
        
        # Map party names and economic responses
        survey_data['mapped_party'] = survey_data['2023 Genel Seçimlerinde hangi partiye oy verdiniz?'].map(self.party_mapping_2023)
        survey_data['econ_future_mapped'] = survey_data[econ_future_col].map(self.econ_future_mapping)
        
        # Filter for negative responses only
        negative_responses = survey_data[survey_data['econ_future_mapped'] == 'Çok Daha Kötü/Daha Kötü']
        
        # Calculate percentages for each party
        party_percentages = {}
        for party in ['AK Parti', 'CHP', 'DEM Parti', 'İYİ Parti', 'MHP']:
            party_total = survey_data[survey_data['mapped_party'] == party]['duzeltilmis_agirlik'].sum()
            party_negative = negative_responses[negative_responses['mapped_party'] == party]['duzeltilmis_agirlik'].sum()
            percentage = (party_negative / party_total * 100) if party_total > 0 else 0
            party_percentages[party] = percentage
        
        # Create or update DataFrame
        if df.empty:
            df = pd.DataFrame(columns=['Months'] + list(party_percentages.keys()))
        
        df.loc[len(df)] = [current_date] + [party_percentages[col] for col in df.columns[1:]]
        return df

    def process_econ_future_age(self, survey_data: pd.DataFrame) -> pd.DataFrame:
        """Process future economic situation breakdown by age (negative responses only)"""
        df = self.read_historical_data('econ_future_age')
        current_date = self.date_formatter.format_date(datetime.now())
        
        # Find the correct column name
        econ_future_col = self._find_column(survey_data, "Önümüzdeki bir yıl içerisinde ekonominin nasıl olacağını düşünüyorsunuz")
        
        # Map economic responses
        survey_data['econ_future_mapped'] = survey_data[econ_future_col].map(self.econ_future_mapping)
        
        # Filter for negative responses only
        negative_responses = survey_data[survey_data['econ_future_mapped'] == 'Çok Daha Kötü/Daha Kötü']
        
        # Calculate percentages for each age group
        age_percentages = {}
        for age_group in ['18-34', '35-54', '55 ve üstü']:
            age_total = survey_data[survey_data['age_group_second'] == age_group]['duzeltilmis_agirlik'].sum()
            age_negative = negative_responses[negative_responses['age_group_second'] == age_group]['duzeltilmis_agirlik'].sum()
            percentage = (age_negative / age_total * 100) if age_total > 0 else 0
            age_percentages[age_group] = percentage
        
        # Create or update DataFrame
        if df.empty:
            df = pd.DataFrame(columns=['Months'] + list(age_percentages.keys()))
        
        df.loc[len(df)] = [current_date] + [age_percentages[col] for col in df.columns[1:]]
        return df

    def calculate_politician_success(self, survey_data: pd.DataFrame, politician_col: str) -> float:
        """Calculate success rate for a single politician"""
        # Get the data for this politician
        politician_data = survey_data[politician_col].copy()
        
        # Remove 'Tanımıyorum' responses
        valid_responses = survey_data[politician_data != 'Tanımıyorum (Anketör Dikkat: Okumayın)']
        
        # Create pivot table
        pivot = pd.pivot_table(
            valid_responses,
            values='duzeltilmis_agirlik',
            index=politician_col,
            aggfunc='sum',
            margins=True
        )
        
        # Calculate percentages
        pivot_pct = pivot / pivot.loc['All'] * 100
        
        
        # Calculate success rate
        success_scores = {}
        for index in pivot.index:
            if index != 'All':
                if index == "1=Çok başarısız":
                    score = 1
                elif index == "10=Çok başarılı":
                    score = 10
                else:
                    try:
                        score = int(index)
                    except:
                        continue
                
                success_scores[score] = pivot_pct.loc[index, 'duzeltilmis_agirlik']
        
        # Calculate weighted average (SUMPRODUCT equivalent)
        success_rate = round(sum(score * (percentage/100) for score, percentage in success_scores.items()), 1)
        
        
        return success_rate

    def process_politician_success(self, survey_data: pd.DataFrame) -> pd.DataFrame:
        """Process current month's politician success rates"""
        # Define the mapping of politicians to their question columns
        politician_columns = {
            'Recep Tayyip Erdoğan': 'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [Recep Tayyip Erdoğan]',
            'Özgür Özel': 'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [Özgür Özel]',
            'Ekrem İmamoğlu': 'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [Ekrem İmamoğlu]',
            'Devlet Bahçeli': 'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [Devlet Bahçeli]',
            'Tülay Hatimoğulları Oruç': 'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [Tülay Hatimoğulları Oruç]',
            'Mansur Yavaş': 'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [Mansur Yavaş]',
            'Mahmut Arıkan': 'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [Mahmut Arıkan]',
            'Muharrem İnce': 'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [Muharrem İnce]',
            'Ümit Özdağ': 'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [Ümit Özdağ]',
            'Erkan Baş': 'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [Erkan Baş]',
            'Fatih Erbakan': 'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [Fatih Erbakan]',
            'Müsavat Dervişoğlu': 'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [Müsavat Dervişoğlu]',
            'Yavuz Ağıralioğlu': 'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [Yavuz Ağıralioğlu]'
        }
        
        # Calculate success rates for all politicians
        success_rates = {}
        for politician, column in politician_columns.items():
            success_rates[politician] = self.calculate_politician_success(survey_data, column)
        
        # Create DataFrame for the chart
        return pd.DataFrame(list(success_rates.items()), columns=['Politician', 'Success Rate'])

    def process_politician_success_main(self, survey_data: pd.DataFrame) -> pd.DataFrame:
        """Process historical data for main politicians' success rates"""
        df = self.read_historical_data('politician_success_main')
        current_date = self.date_formatter.format_date(datetime.now())
        
        # Replace 0s with None in historical data
        if not df.empty:
            df = df.replace(0, None)
        
        # Calculate success rates for main politicians
        main_politicians = ['Recep Tayyip Erdoğan', 'Özgür Özel', 'Devlet Bahçeli', 
                          'Ekrem İmamoğlu', 'Mansur Yavaş', 'Fatih Erbakan']
        
        success_rates = {}
        for politician in main_politicians:
            column = f'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [{politician}]'
            rate = self.calculate_politician_success(survey_data, column)
            success_rates[politician] = rate if rate > 0 else None
        
        # Create or update DataFrame
        if df.empty:
            df = pd.DataFrame(columns=['Months'] + main_politicians)
        
        df.loc[len(df)] = [current_date] + [success_rates[col] for col in df.columns[1:]]
        return df

    def process_politician_success_second(self, survey_data: pd.DataFrame) -> pd.DataFrame:
        """Process historical data for secondary politicians' success rates"""
        df = self.read_historical_data('politician_success_second')
        current_date = self.date_formatter.format_date(datetime.now())
        
        # Replace 0s with None in historical data
        if not df.empty:
            df = df.replace(0, None)
        
        # Calculate success rates for secondary politicians
        second_politicians = ['Muharrem İnce', 'Erkan Baş', 'Ümit Özdağ', 'Müsavat Dervişoğlu',
                            'Tülay Hatimoğulları Oruç', 'Yavuz Ağıralioğlu', 'Mahmut Arıkan']
        
        success_rates = {}
        for politician in second_politicians:
            column = f'Sayacağım siyasetçileri 1-10 arası ne kadar başarılı buluyorsunuz? Lütfen tanımadığınız siyasetçi olursa belirtiniz. (1=Çok başarısız, 10=Çok başarılı) [{politician}]'
            rate = self.calculate_politician_success(survey_data, column)
            success_rates[politician] = rate if rate > 0 else None
        
        # Create or update DataFrame
        if df.empty:
            df = pd.DataFrame(columns=['Months'] + second_politicians)
        
        df.loc[len(df)] = [current_date] + [success_rates[col] for col in df.columns[1:]]
        return df

    def process_subsistence(self, survey_data: pd.DataFrame) -> pd.DataFrame:
        """Process main subsistence data"""
        df = self.read_historical_data('subsistence')
        current_date = self.date_formatter.format_date(datetime.now())
        
        # Find the correct column name
        subsistence_col = self._find_column(survey_data, "Aşağıdaki sayılan ifadelerden hangisine katılırsınız")
        
        # Define the order of responses in the historical data
        responses = [
            'Gelirim giderimi karşılamadı.',
            'Gelirim giderimi ucu ucuna karşıladı.',
            'Gelirim giderlerimin üzerinde oldu.',
            'Gelirim giderlerimi fazlasıyla karşıladı.'
        ]
        
        # Map responses to match exactly
        response_mapping = {
            'Geçtiğimiz ay gelirim giderlerimi karşılamadı.': 'Gelirim giderimi karşılamadı.',
            'Geçtiğimiz ay gelirim giderlerimi ucu ucuna karşıladı.': 'Gelirim giderimi ucu ucuna karşıladı.',
            'Geçtiğimiz ay gelirim giderlerimin üzerinde oldu.': 'Gelirim giderlerimin üzerinde oldu.',
            'Geçtiğimiz ay gelirim giderlerimi fazlasıyla karşıladı.': 'Gelirim giderlerimi fazlasıyla karşıladı.'
        }
        
        survey_data['mapped_subsistence'] = survey_data[subsistence_col].map(response_mapping)
        
        # Calculate total weight
        total_weight = survey_data['duzeltilmis_agirlik'].sum()
        
        # Calculate percentages for each response (row percentages)
        percentages = {}
        for response in responses:
            response_data = survey_data[survey_data['mapped_subsistence'] == response]
            percentage = (response_data['duzeltilmis_agirlik'].sum() / total_weight * 100) if total_weight > 0 else 0
            percentages[response] = percentage
        
        # Create or update DataFrame
        if df.empty:
            df = pd.DataFrame(columns=['Months'] + responses)
        
        df.loc[len(df)] = [current_date] + [percentages[col] for col in df.columns[1:]]
        return df

    def process_subsistence_party(self, survey_data: pd.DataFrame) -> pd.DataFrame:
        """Process subsistence data by party (negative responses only)"""
        df = self.read_historical_data('subsistence_party')
        current_date = self.date_formatter.format_date(datetime.now())
        
        # Find the correct column name
        subsistence_col = self._find_column(survey_data, "Aşağıdaki sayılan ifadelerden hangisine katılırsınız")
        
        # Map party names
        survey_data['mapped_party'] = survey_data['2023 Genel Seçimlerinde hangi partiye oy verdiniz?'].map(self.party_mapping_2023)
        
        # Map responses to match exactly
        response_mapping = {
            'Geçtiğimiz ay gelirim giderlerimi karşılamadı.': 'Gelirim giderimi karşılamadı.',
            'Geçtiğimiz ay gelirim giderlerimi ucu ucuna karşıladı.': 'Gelirim giderimi ucu ucuna karşıladı.',
            'Geçtiğimiz ay gelirim giderlerimin üzerinde oldu.': 'Gelirim giderlerimin üzerinde oldu.',
            'Geçtiğimiz ay gelirim giderlerimi fazlasıyla karşıladı.': 'Gelirim giderlerimi fazlasıyla karşıladı.'
        }
        
        survey_data['mapped_subsistence'] = survey_data[subsistence_col].map(response_mapping)
        
        # Filter for negative responses
        negative_responses = survey_data[
            (survey_data['mapped_subsistence'] == 'Gelirim giderimi karşılamadı.') |
            (survey_data['mapped_subsistence'] == 'Gelirim giderimi ucu ucuna karşıladı.')
        ]
        
        # Calculate percentages for each party
        party_percentages = {}
        for party in ['AK Parti', 'CHP', 'DEM Parti', 'İYİ Parti', 'MHP']:
            party_total = survey_data[survey_data['mapped_party'] == party]['duzeltilmis_agirlik'].sum()
            party_negative = negative_responses[negative_responses['mapped_party'] == party]['duzeltilmis_agirlik'].sum()
            percentage = (party_negative / party_total * 100) if party_total > 0 else 0
            party_percentages[party] = percentage
        
        # Create or update DataFrame
        if df.empty:
            df = pd.DataFrame(columns=['Months'] + list(party_percentages.keys()))
        
        df.loc[len(df)] = [current_date] + [party_percentages[col] for col in df.columns[1:]]
        return df

    def save_updated_data(self, data, sheet_name: str = 'party_votes'):
        """Save updated data to the specified sheet"""
        try:
            if isinstance(data, dict):
                # For education and age breakdowns
                with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    for sheet, df in data.items():
                        df = df.replace([np.inf, -np.inf], np.nan).fillna(value=np.nan)
                        df.to_excel(writer, sheet_name=sheet, index=False)
            else:
                # For single dataframe updates
                with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    data = data.replace([np.inf, -np.inf], np.nan).fillna(value=np.nan)
                    data.to_excel(writer, sheet_name=sheet_name, index=False)
                
        except Exception as e:
            print(f"Error saving updated data to sheet {sheet_name}: {str(e)}")
            if "No such file or directory" in str(e):
                if isinstance(data, dict):
                    with pd.ExcelWriter(self.file_path) as writer:
                        for sheet, df in data.items():
                            df = df.replace([np.inf, -np.inf], np.nan).fillna(value=np.nan)
                            df.to_excel(writer, sheet_name=sheet, index=False)
                else:
                    data = data.replace([np.inf, -np.inf], np.nan).fillna(value=np.nan)
                    data.to_excel(self.file_path, sheet_name=sheet_name, index=False)