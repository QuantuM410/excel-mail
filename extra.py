import pandas as pd
from faker import Faker

# Create Faker instance
fake = Faker()

# Generate random data
data = []
email = 'kartikeys410@gmail.com'
for _ in range(10):  # Number of rows to generate
    name = fake.name()
    grade = fake.random_int(min=1, max=12)

    data.append([name, grade, email])

# Create DataFrame
df = pd.DataFrame(data, columns=['Name', 'Grade', 'Email'])

# Save DataFrame to Excel file
df.to_excel('sample_excel_mail.xlsx', index=False)
