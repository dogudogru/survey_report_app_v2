import pandas as pd
from typing import Dict
from config.constants import PARTY_MAPPING  # Changed to absolute import

class DataProcessor:
    @staticmethod
    def process_survey_data(df: pd.DataFrame, question_column: str, weight_column: str = 'duzeltilmis_agirlik') -> Dict[str, float]:
        """Process survey data and calculate weighted percentages"""
        # Map party names
        df['Party'] = df[question_column].map(
            lambda x: PARTY_MAPPING.get(x, 'Diğer')
        )
        
        # Calculate weighted percentages
        total_weight = df[weight_column].sum()
        party_results = df.groupby('Party')[weight_column].sum().reset_index()
        party_results['Percentage'] = (party_results[weight_column] / total_weight) * 100
        
        # Round percentages to one decimal place
        party_results['Percentage'] = party_results['Percentage'].round(1)
        
        return {row['Party']: row['Percentage'] for _, row in party_results.iterrows()}

    @staticmethod
    def prepare_sorted_data(party_data: pd.DataFrame) -> pd.DataFrame:
        """Prepare sorted data with 'Diğer' always at the bottom"""
        diger_data = party_data[party_data['Party'] == 'Diğer']
        other_parties = party_data[party_data['Party'] != 'Diğer']
        sorted_parties = other_parties.sort_values('Percentage', ascending=False)
        if not diger_data.empty:
            sorted_parties = pd.concat([sorted_parties, diger_data])
        return sorted_parties
