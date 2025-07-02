# AWS Comprehensive Resource Extractor
A robust Python tool to extract detailed information from 25+ AWS services and export everything to a professionally formatted Excel workbook.
Supports manual AWS profile selection, multi-region support, and is ideal for security reviews, audits, and cloud posture assessments.

Service list:
Ec2 
QuickSight
Elastic Load Balancing
Elastic Container Service
DocumentDB
Elastic Container Service for Kubernetes
Textract
Rekognition
Key Management Service
Bedrock
CloudTrail
Systems Manager
GuardDuty
API Gateway
Data Transfer
CloudFormation
Location Service
Neptune
Service Catalog
AWS Glue 
CodeArtifact
Dynamo DB 
Memory DB
VPC 
Sage  Maker 
Elastic File System
OpenSearch Service
Security hub 
Elastic Container Service for Kubernetes
X-Ray
EC2-Other


# 🚦 Prerequisites
Python 3.8+

Make sure you have Python 3.8 or newer installed.
Check your version:

python --version

Virtual Environment (Recommended)
Create and activate a virtual environment to keep dependencies isolated:

python -m venv venv

On Windows:
venv\Scripts\activate

On macOS/Linux:
source venv/bin/activate

#AWS CLI Configured

You must have the AWS CLI installed and at least one profile configured:

aws configure --profile my-aws-profile

The script uses your AWS CLI credentials and profiles for authentication.
📦 Installation

Clone the repository:

git clone https://github.com/1o1swapnil/aws-comprehensive-extractor.git

cd aws-comprehensive-extractor

Install dependencies:

pip install boto3 pandas openpyxl

⚙️ Usage

# With a specific AWS profile and region
python aws_comprehensive_extractor.py --profile my-aws-profile --region us-west-2

# With default AWS credentials
python aws_comprehensive_extractor.py --region eu-west-1

The script will generate an Excel file like AWS_Comprehensive_Report_YYYYMMDD_HHMMSS.xlsx in the current directory.
Each AWS service will have its own sheet, and a summary sheet will provide an overview.

📝 Example Output
Summary Sheet:

Service Sheet (e.g., EC2):

🔒 Permissions

You must have sufficient IAM permissions for each AWS service you wish to extract data from.
The script will skip services where access is denied and continue with the rest.

🛠️ Contributing

Contributions are welcome!
Feel free to open issues or submit pull requests for new features, bug fixes, or service support.

Fork the repo

Create your feature branch (git checkout -b feature/AmazingFeature)

Commit your changes (git commit -m 'Add some AmazingFeature')

Push to the branch (git push origin feature/AmazingFeature)

Open a pull request
📄 License
MIT License. See LICENSE for details.

🙋 FAQ
Q: What if I get access denied errors?
A: The script will log the error and continue. Make sure your AWS profile has the necessary permissions.

Q: Can I add more services?
A: Yes! The script is modular. Add your extraction logic and register it in the main extraction list.

🤝 Contact
For questions, suggestions, or support, open an issue or contact 1o1swapnilk@gmail.com.

Star this repo if you find it useful! ⭐
