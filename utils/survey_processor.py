# utils/survey_processor.py
from typing import Dict, List
import pandas as pd

class SurveyProcessor:
    def __init__(self, questions_config: Dict[str, Dict]):
        """
        Initialize with questions configuration
        
        questions_config format:
        {
            'question_column_name': {
                'type': 'single_choice|multiple_choice',
                'mapping': {'old_value': 'new_value', ...},
                'weight_column': 'column_name'
            },
            ...
        }
        """
        self.questions_config = questions_config

    def process_all_questions(self, df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
        """Process all configured questions and return their pivot tables"""
        results = {}
        
        for question, config in self.questions_config.items():
            if config['type'] == 'single_choice':
                results[question] = self._process_single_choice(
                    df,
                    question,
                    config.get('mapping', {}),
                    config.get('weight_column', 'duzeltilmis_agirlik')
                )
            elif config['type'] == 'multiple_choice':
                results[question] = self._process_multiple_choice(
                    df,
                    question,
                    config.get('mapping', {}),
                    config.get('weight_column', 'duzeltilmis_agirlik')
                )
                
        return results

    def _process_single_choice(
        self,
        df: pd.DataFrame,
        question: str,
        mapping: Dict[str, str],
        weight_column: str
    ) -> pd.DataFrame:
        """Process single choice questions"""
        # Apply mapping if provided
        if mapping:
            df = df.copy()
            df['Response'] = df[question].map(lambda x: mapping.get(x, x))
        else:
            df['Response'] = df[question]

        # Calculate weighted percentages
        total_weight = df[weight_column].sum()
        results = df.groupby('Response')[weight_column].sum().reset_index()
        results['Percentage'] = (results[weight_column] / total_weight) * 100
        results['Percentage'] = results['Percentage'].round(1)

        return results[['Response', 'Percentage']]

    def _process_multiple_choice(
        self,
        df: pd.DataFrame,
        question: str,
        mapping: Dict[str, str],
        weight_column: str
    ) -> pd.DataFrame:
        """Process multiple choice questions"""
        # Split multiple responses and process each option
        responses = df[question].str.get_dummies(sep=';')
        
        # Apply weights
        weighted_responses = responses.multiply(df[weight_column], axis=0)
        
        # Calculate percentages
        total_weight = df[weight_column].sum()
        percentages = (weighted_responses.sum() / total_weight * 100).round(1)
        
        # Apply mapping if provided
        if mapping:
            percentages.index = percentages.index.map(lambda x: mapping.get(x, x))
            
        return pd.DataFrame({
            'Response': percentages.index,
            'Percentage': percentages.values
        })