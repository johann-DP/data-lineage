import uuid
import pandas as pd
import numpy as np
from faker import Faker

fake = Faker()

class DataGenerator:
    def __init__(self, n_rows=1000):
        self.n_rows = n_rows
        self.fake = Faker()
    
    #@staticmethod
    def generate_ids(self, n_rows=1000):
        return [str(uuid.uuid4()) for _ in range(n_rows)]
    
    #@staticmethod
    def generate_unique_ids(self, n):
       return [str(uuid.uuid4()) for _ in range(n)]

    def generate_base_params(self):
        """Génère un DataFrame contenant des données factices pour les paramètres de base."""
        df = pd.DataFrame({
            'ID': self.generate_ids(),
            'Date': pd.date_range(start='1/1/2020', periods=self.n_rows),
            'Interest_Rate': np.random.uniform(low=0.01, high=0.05, size=self.n_rows),
            'Inflation_Rate': np.random.uniform(low=0.01, high=0.05, size=self.n_rows)
        })
        return df

    def generate_demographic_data(self):
        """Génère un DataFrame contenant des données factices pour les données démographiques."""
        df = pd.DataFrame({
            'ID': self.generate_ids(),
            'Age': np.random.randint(low=18, high=70, size=self.n_rows),
            'Sex': np.random.choice(['Male', 'Female'], size=self.n_rows),
            'Occupation': [self.fake.job() for _ in range(self.n_rows)]
        })
        return df

    def generate_claim_data(self):
        """Génère un DataFrame contenant des données factices pour les données de sinistre."""
        df = pd.DataFrame({
            'ID': self.generate_ids(),
            'Claim_Date': [self.fake.date_between(start_date='-5y', end_date='today') for _ in range(self.n_rows)],
            'Claim_Type': np.random.choice(['Collision', 'Fire', 'Theft', 'Natural Disaster'], size=self.n_rows),
            'Claim_Amount': np.random.uniform(low=1000, high=100000, size=self.n_rows)
        })
        return df

    def generate_premium_data(self):
        """Génère un DataFrame contenant des données factices pour les données de prime."""
        df = pd.DataFrame({
            'ID': self.generate_ids(),
            'Premium_Date': [self.fake.date_between(start_date='-3y', end_date='today') for _ in range(self.n_rows)],
            'Premium_Amount': np.random.uniform(low=300, high=3000, size=self.n_rows),
            'Contract_Type': np.random.choice(['Life', 'Health', 'Auto', 'Home'], size=self.n_rows),
        })
        return df

    def generate_contract_data(self):
        """Génère un DataFrame contenant des données factices pour les données des contrats."""
        df = pd.DataFrame({
            'Contract_ID': self.generate_ids(),
            'Contract_Date': [self.fake.date_between(start_date='-3y', end_date='today') for _ in range(self.n_rows)],
            'Contract_Amount': np.random.uniform(low=3000, high=30000, size=self.n_rows),
            'Contract_Type': np.random.choice(['Life', 'Health', 'Auto', 'Home'], size=self.n_rows),
        })
        return df

    def generate_claim_history_data(self):
        """Génère un DataFrame contenant des données factices pour l'historique des sinistres."""
        df = pd.DataFrame({
            'Claim_ID': self.generate_ids(),
            'Contract_ID': self.generate_ids(),  # Assume each claim is related to a unique contract
            'Claim_Date': [self.fake.date_between(start_date='-5y', end_date='today') for _ in range(self.n_rows)],
            'Claim_Type': np.random.choice(['Collision', 'Fire', 'Theft', 'Natural Disaster'], size=self.n_rows),
            'Claim_Amount': np.random.uniform(low=1000, high=100000, size=self.n_rows)
        })
        return df

    def generate_actuarial_params(self, i, n_rows=3):  # Only three rows needed for these parameters
        """Génère un DataFrame contenant des données factices pour les paramètres actuariels."""
        df = pd.DataFrame({
            'Parameter_Name': ['Discount Rate', 'Loss Development Factor', 'Loss Adjustment Expense Ratio'],
            'Parameter_Value': [0.05 + i*0.01, 1.2 + i*0.1, 0.15 + i*0.02],
        })
        return df

    def generate_stress_scenarios(self, n_rows=200):
        df = pd.DataFrame({
            'Scenario_ID': self.generate_ids(n_rows),
            'Scenario_Name': ['Scenario ' + str(i) for i in range(1, n_rows + 1)],
            'Interest_Rate_Shock': np.random.uniform(low=-0.05, high=0.05, size=n_rows),
            'Equity_Price_Shock': np.random.uniform(low=-0.2, high=0.2, size=n_rows),
        })
        return df

    def generate_risk_params(self, n_rows=3):
        df = pd.DataFrame({
            'Parameter_Name': ['Interest Rate Volatility', 'Equity Price Volatility', 'Correlation'],
            'Parameter_Value': [0.01, 0.2, 0.5],
        })
        return df

    def generate_risk_exposures(self, n_rows=1000):
        df = pd.DataFrame({
            'Asset_ID': self.generate_ids(n_rows),
            'Asset_Value': np.random.uniform(low=10000, high=1000000, size=n_rows),
            'Asset_Type': np.random.choice(['Bond', 'Equity', 'Real Estate', 'Cash'], size=n_rows),
        })
        return df

    def generate_ids_2(self, n=2000):
        return ['ID' + str(i) for i in range(1, n+1)]

    def generate_asset_data(self, n_rows=2000):  
        df = pd.DataFrame({
            'Asset_ID': self.generate_ids_2(n_rows), 
            'Asset_Value': np.random.uniform(low=10000, high=1000000, size=n_rows),
            'Asset_Type': np.random.choice(['Bond', 'Equity', 'Real Estate', 'Cash'], size=n_rows),
        })
        return df

    def generate_liability_data(self, n_rows=2000):  
        df = pd.DataFrame({
            'Liability_ID': self.generate_ids_2(n_rows),  
            'Liability_Value': np.random.uniform(low=10000, high=1000000, size=n_rows),
            'Liability_Type': np.random.choice(['Reserve', 'Debt', 'Other'], size=n_rows),
        })
        return df

    def generate_regulatory_data(self, n_rows=3):
        df = pd.DataFrame({
            'Parameter_Name': ['Solvency Ratio', 'Minimum Capital Requirement', 'Risk Margin'],
            'Parameter_Value': [1.5, 1000000, 0.05],
        })
        return df

    def generate_model_parameters(self, n_rows=6):  
        df = pd.DataFrame({
            'Parameter_Name': ['Discount Rate', 'Risk Free Rate', 'Volatility', 'Inflation', 'Growth Rate', 'Dividend Yield'],  
            'Parameter_Value': np.random.uniform(low=0.01, high=0.05, size=n_rows),
        })
        return df

    def generate_model_results(self, ids):  
        df = pd.DataFrame({
            'Scenario_ID': ids,  
            'Scenario_Result': np.random.uniform(low=-0.2, high=0.2, size=len(ids)),
        })
        return df

    def generate_monte_carlo_data(self, ids):  
        df = pd.DataFrame({
            'Simulation_ID': ids,
            'Simulation_Result': np.random.normal(loc=0, scale=0.05, size=len(ids)),
        })
        return df

    def generate_reinsurance_contracts(self):
        df = pd.DataFrame({
            'ID': self.generate_ids(),
            'Contract_ID': ['Contract ' + str(i) for i in range(1, self.n_rows+1)],
            'Contract_Start_Date': [self.fake.date_between(start_date='-3y', end_date='today') for _ in range(self.n_rows)],
            'Contract_End_Date': [self.fake.date_between(start_date='today', end_date='+3y') for _ in range(self.n_rows)],
            'Contract_Premium': np.random.uniform(low=50000, high=500000, size=self.n_rows),
        })
        return df

    def generate_claim_history(self, contract_ids):
        df = pd.DataFrame({
            'ID': self.generate_ids(),
            'Contract_ID': np.random.choice(contract_ids, size=self.n_rows),  # Reference to 'Contract_ID' of 'Reinsurance_Contracts'
            'Claim_Date': [self.fake.date_between(start_date='-5y', end_date='today') for _ in range(self.n_rows)],
            'Claim_Amount': np.random.uniform(low=50000, high=5000000, size=self.n_rows),
            'Claim_Type': np.random.choice(['Property Damage', 'Bodily Injury', 'Product Liability', 'Others'], size=self.n_rows),
        })
        return df

    def generate_claim_details(self, n_rows=1000):
        df = pd.DataFrame({
            'ID': self.generate_ids(),
            'Claim_ID': ['Claim ' + str(i) for i in range(1, n_rows+1)],
            'Claim_Amount': np.random.uniform(low=500, high=5000, size=n_rows),
            'Claim_Type': np.random.choice(['Accidental', 'Natural', 'Theft'], size=n_rows),
        })
        return df

    def generate_reinsurance_pricing_params(self):
        df = pd.DataFrame({
            'Parameter_Name': ['Discount Rate', 'Expected Loss Ratio', 'Risk Margin'],
            'Parameter_Value': np.random.uniform(low=0.01, high=0.05, size=3),
        })
        return df

    def generate_market_indices(self):
        df = pd.DataFrame({
            'ID': self.generate_ids(),
            'Index_Name': ['Index ' + str(i) for i in range(1, self.n_rows+1)],
            'Index_Value': np.random.uniform(low=1000, high=10000, size=self.n_rows),
        })
        return df

    def generate_customer_details(self):
        df = pd.DataFrame({
            'ID': self.generate_ids(),
            'Customer_Name': [self.fake.name() for _ in range(self.n_rows)],
            'Customer_Age': np.random.randint(low=20, high=80, size=self.n_rows),
            'Customer_Gender': np.random.choice(['Male', 'Female'], size=self.n_rows),
        })
        return df

    def generate_policy_details(self):
        df = pd.DataFrame({
            'ID': self.generate_ids(),
            'Policy_ID': ['Policy ' + str(i) for i in range(1, self.n_rows+1)],
            'Date_of_Purchase': [self.fake.date_between(start_date='-5y', end_date='today') for _ in range(self.n_rows)],
            'Customer_ID': self.generate_ids(),
            'Premium': np.random.uniform(low=500, high=5000, size=self.n_rows),
            'Type_of_Cover': np.random.choice(['Comprehensive', 'Third Party'], size=self.n_rows),
        })
        return df

    def generate_agent_details(self, n_rows=1000):
        df = pd.DataFrame({
            'ID': self.generate_ids(),
            'Agent_Name': [fake.name() for _ in range(n_rows)],
            'Agent_Commission': np.random.uniform(low=5, high=20, size=n_rows),
        })
        return df

    def generate_product_details(self, n_rows=1000):
        df = pd.DataFrame({
            'ID': self.generate_ids(),
            'Product_Name': ['Product ' + str(i) for i in range(1, n_rows+1)],
            'Product_Price': np.random.uniform(low=100, high=1000, size=n_rows),
        })
        return df
    