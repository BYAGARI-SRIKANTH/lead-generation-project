import pandas as pd

# Lead data
data = [
    ["TechCorp Solutions", "contact@techcorp.com", "www.techcorp.com", "Hyderabad"],
    ["DataMinds AI", "info@dataminds.ai", "www.dataminds.ai", "Bangalore"],
    ["GreenEarth NGO", "hello@greenearth.org", "www.greenearth.org", "Delhi"],
    ["CloudBase Systems", "support@cloudbase.in", "www.cloudbase.in", "Mumbai"],
    ["HealthFirst India", "contact@healthfirst.in", "www.healthfirst.in", "Chennai"],
    ["EduLearn Platform", "team@edulearn.com", "www.edulearn.com", "Pune"],
    ["FinTrack Solutions", "info@fintrack.co", "www.fintrack.co", "Hyderabad"],
    ["SmartCity Labs", "hello@smartcity.in", "www.smartcity.in", "Delhi"],
    ["AgriTech Ventures", "contact@agritech.org", "www.agritech.org", "Ahmedabad"],
    ["SecureNet India", "info@securenet.in", "www.securenet.in", "Mumbai"],
    ["MediaStream India", "", "www.mediastream.in", "Mumbai"],
    ["Robotics India", "", "www.robotics.in", "Chennai"],
    ["BlockChain Co", "info@blockchain.co", "www.blockchain.co", "Bangalore"],
    ["VR World India", "", "www.vrworld.in", "Hyderabad"],
    ["SupplyChain Pro", "", "www.supplychain.in", "Pune"],
    ["CleanWater NGO", "", "www.cleanwater.org", "Delhi"],
    ["InsureTech India", "team@insuretech.in", "www.insuretech.in", "Mumbai"],
    ["GameZone Studios", "", "www.gamezone.in", "Bangalore"],
    ["LegalTech Hub", "", "www.legaltech.in", "Delhi"],
    ["TravelNest India", "", "www.travelnest.in", "Goa"],
    ["FoodBox Solutions", "", "www.foodbox.in", "Hyderabad"],
    ["FarmFresh Tech", "", "www.farmfresh.in", "Ahmedabad"],
    ["EnergyPlus India", "", "www.energyplus.in", "Chennai"],
    ["NextGen AI Labs", "", "www.nextgenai.in", "Bangalore"],
    ["UrbanBuild Systems", "", "www.urbanbuild.in", "Mumbai"],

    # Duplicate row
    ["TechCorp Solutions", "contact@techcorp.com", "www.techcorp.com", "Hyderabad"],

    # Missing email rows
    ["BioGen Research", "", "www.biogen.in", "Hyderabad"],
    ["RetailPro India", "", "www.retailpro.in", "Delhi"],
    ["AutoDrive Tech", "", "www.autodrive.co", "Pune"]
    
]

# Create DataFrame
df = pd.DataFrame(data, columns=["Name", "Email", "Website", "Location"])

# -------------------------
# Data Cleaning
# -------------------------

# Remove duplicates
df.drop_duplicates(inplace=True)
df.reset_index(drop=True, inplace=True)

# Fill missing emails automatically
def generate_email(website):
    domain = website.replace("www.", "").strip()
    return "info@" + domain

df["Email"] = df.apply(
    lambda row: generate_email(row["Website"]) if row["Email"] == "" else row["Email"],
    axis=1
)

# Remove extra spaces
df["Name"] = df["Name"].str.strip()
df["Location"] = df["Location"].str.strip()

# -------------------------
# Save to Excel
# -------------------------
df.to_excel("leads_output.xlsx", index=False)

# Show output
print("Excel file created successfully!")
print(df)