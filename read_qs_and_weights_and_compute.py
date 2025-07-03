#%% ============================= Load structure & weights from Excel

import pandas as pd

# filename = r'C:\Users\CG14739\OneDrive - Cerved\Python Scripts\ESGRecode\Struttura_e_pesi.xlsx'
# values_df = pd.read_excel(filename, sheet_name='KPI - Values', header=None)

url       = "https://raw.githubusercontent.com/marcogalbiati/ESGRecode/main/Data/Struttura_e_pesi.xlsx"
values_df = pd.read_excel(url, sheet_name='KPI - Values', header=None)

# Find all dimension rows (e.g., ENV.POL, SOC.WORKFORCE, etc.) and their weights
dimension_mask = values_df[0].astype(str).str.match(r'^[A-Z]+\.[A-Z]+$')
dimensions = values_df[dimension_mask][[0, 1, 4]].copy()
dimensions.columns = ['Dimension', 'Dimension_Name', 'Weight']
dimensions['Weight'] = pd.to_numeric(dimensions['Weight'], errors='coerce')

# For each Dimension, collect all questions under it
question_rows = []
for idx, row in dimensions.iterrows():
    dim_code = row['Dimension']
    # Questions are in the following rows, where column 1 matches e.g. ENV.1, ENV.2, etc.
    i = idx + 1
    while i < len(values_df):
        code = str(values_df.iloc[i, 1])
        if pd.isna(code) or not code.startswith(dim_code.split('.')[0]):
            break
        if '.' in code and any(c.isdigit() for c in code.split('.')[-1]):
            question_rows.append({
                'Dimension': dim_code,
                'Question_Code': code,
                'Question_Text': values_df.iloc[i, 2],
                'Dimension_Weight': row['Weight']
            })
        i += 1
questions_df = pd.DataFrame(question_rows)

# Add 'Pillar' column by splitting the 'Dimension' string at the dot
questions_df['Pillar'] = questions_df['Dimension'].str.split('.').str[0]

# Tidy up the columns
desired_order = [
    'Pillar',
    'Dimension',
    'Question_Code',
    'Question_Text',
    'Dimension_Weight'
]
questions_df = questions_df.reset_index(drop=True)
questions_df = questions_df[desired_order]

# # Compute pillar weights. A Pillar's weight is the sum of the Weights of the Dimensions in the Pillar
# unique_combinations    = questions_df.drop_duplicates(subset=['Dimension_Weight', 'Pillar'])
# sum_weights_per_pillar = unique_combinations.groupby('Pillar', as_index=False)['Dimension_Weight'].sum()
# sum_weights_per_pillar = sum_weights_per_pillar.rename(columns={'Dimension_Weight': 'Pillar_weight'})
# questions_df = questions_df.merge(sum_weights_per_pillar, on='Pillar', how='left')

# ===  THE LOGIC of the scoring system IS:
# WEIGHTS ARE SET FOR THE 3 PILLARS, DOWNSTREAM EQUAL SPLIT (WEIGHTS DEPEND ON N. OF ITEMS AT EACH LEVEL)
# EQUIVALENTLY: WEIGHTS ARE SET AT D-EVEL (EQUAL FOR EACH D IN A PILLAR),
# DOWNSTREAM EQUAL SPLIT,
# UPSTREAM RULE FIXED BY SAID WEIGHTS AND N. OF DIMENSIONS


#%% ============================= COMPUTE SCORE, STARTING FROM MATRIX OF QUESTION SCORES

# Load q-scores of firms
url       = "https://raw.githubusercontent.com/marcogalbiati/ESGRecode/main/Data/Firms_question_scores_anon.xlsx"
scores_df = pd.read_excel(url, sheet_name='KPI - Scores', header=None)


#%% Pick one firm & score it
firm_col = 9    # Pick a column ie a Company  (5 is firm A)
compname  = scores_df.iloc[0,firm_col]

# Extract the relevant columns: question code and the chosen company's score
scores_to_merge = scores_df[[1, firm_col]].copy()
scores_to_merge.columns = ['Question_Code', 'Score']

# Join with qustions_df
questions_df_scored = questions_df.copy()
questions_df_scored = questions_df_scored.merge(
    scores_to_merge,
    on='Question_Code',
    how='left'
)

# Function to compute dimension averages (fill with zero / ignore options):
def compute_dimension_scores(questions_df, fillna_zero=False):
    """
    Compute dimension averages.
    If fillna_zero is True, treat missing question scores as 0.
    If False, ignore missing scores in the average.
    """
    df = questions_df.copy()
    if fillna_zero:
        # Fill missing scores with 0
        df['Score'] = df['Score'].fillna(0)
        # Average is over all questions (including those with no score, now 0)
        dimension_scores = (
            df.groupby(['Dimension', 'Dimension_Weight'])
            .agg(Dimension_Avg=('Score', 'mean'))
            .reset_index()
        )
    else:
        # Drop missing scores, average only over available scores
        dimension_scores = (
            df.dropna(subset=['Score'])
            .groupby(['Dimension', 'Dimension_Weight'])
            .agg(Dimension_Avg=('Score', 'mean'))
            .reset_index()
        )
    return dimension_scores

# Execute above function
dimension_scores = compute_dimension_scores(questions_df_scored, fillna_zero=False)  # or False

# Function to compute total ESG score:
def compute_esg_score(dimension_scores):
    dimension_scores['Weighted'] = dimension_scores['Dimension_Avg'] * dimension_scores['Dimension_Weight']
    return dimension_scores['Weighted'].sum() / dimension_scores['Dimension_Weight'].sum()

# Execute above function
esg_score        = compute_esg_score(dimension_scores)
print(f"ESG Score of {compname} (col. {firm_col:.0f}): {esg_score:.2f}")

# # Pick q-scores of a firm
# firm_col = 125     # Pick a col ie a firm
# firmname  = scores_df.iloc[0,firm_col]

# # Build a mapping from question code to score
# score_map = {}
# for i in range(len(scores_df)):
#     code = str(scores_df.iloc[i, 1])
#     if code in questions_df['Question_Code'].values:
#         score = scores_df.iloc[i, firm_col]
#         try:
#             score = float(score)
#         except:
#             continue
#         score_map[code] = score

# # Merge scores into questions_df
# questions_df_scored['Score'] = questions_df_scored['Question_Code'].map(score_map)

# # Function to compute ESG score
# def calculate_esg_score(questions_df_scored):
#     """
#     Calculate dimension averages and weighted ESG score from questions_df_scored.
#     """
#     dimension_scores = (
#         questions_df_scored
#         .dropna(subset=['Score'])
#         .groupby(['Dimension'])
#         .agg(Dimension_Avg=('Score', 'mean'), Dimension_Weight=('Dimension_Weight', 'first'))
#         .reset_index()
#     )
#     dimension_scores['Weighted'] = dimension_scores['Dimension_Avg'] * dimension_scores['Dimension_Weight']
#     esg_score = dimension_scores['Weighted'].sum() / dimension_scores['Dimension_Weight'].sum()
#     return esg_score

# z = calculate_esg_score(questions_df_scored)
# # print(f"ESG Score for {firmname}: {z:.2f}")