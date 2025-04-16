"""Configuration settings for policy data generation"""

# Product specifications
PRODUCT_NAME = "Secure20 Term Life"
ENTRY_AGE = 30
TERMS = [5, 10, 15]
BASE_SA = 10000
PREMIUM_RATE = 0.05  # 5% of SA
DEFAULT_NUM_POLICIES = 1000  # Set default to 1000 policies
RANDOM_SEED = 42

# Policy status probabilities
STATUS_PROBABILITIES = {
    'In Force': 0.95,
    'Death Claim': 0.05
}

# Date settings
PURCHASE_DATE_RANGE_YEARS = 2  # Generate purchase dates within last 2 years