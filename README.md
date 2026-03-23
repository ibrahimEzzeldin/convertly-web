# Convertly - File Conversion Service

A modern web application for converting files between different formats with multilingual support and payment integration.

## Features

- **File Format Conversion** - Convert documents, images, and other files between multiple formats
- **Multilingual Support** - User interface available in multiple languages with ISO language codes
- **Translation Service** - Built-in translation capabilities for converted content
- **Freemium Model** - Free tier with limited conversions, premium tier for unlimited access
- **PayPal Integration** - Secure payment processing for premium features
- **Voucher System** - Support for promotional codes and vouchers
- **User-Friendly Interface** - Clean, responsive web interface for easy file uploads
- **SEO Optimized** - Sitemap support and search engine optimization
- **Secure** - Content Security Policy headers and secure session management
- **Containerized Deployment** - Docker support for easy deployment to cloud platforms

## Tech Stack

- **Backend**: Python Flask
- **Frontend**: HTML, CSS, JavaScript
- **File Processing**: LibreOffice (installed via Docker)
- **Deployment**: Docker, Render.yaml
- **APIs**: Anthropic Claude API, PayPal API
- **Database**: Session-based storage

## Prerequisites

- Python 3.8+
- Docker (for containerized deployment)
- LibreOffice (for file conversion)
- API Keys:
  - Anthropic Claude API key
  - PayPal Client ID and Secret

## Installation

### Local Development

1. Clone the repository:
```bash
git clone https://github.com/ibrahimEzzeldin/convertly-web.git
cd convertly-web
```

2. Create a virtual environment:
```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Create `.env` file with your configuration:
```bash
cp .env.example .env
# Edit .env with your API keys and settings
```

5. Run the application:
```bash
python run.py
```

The application will be available at `http://localhost:5000`

### Docker Deployment

Build and run with Docker:
```bash
docker build -t convertly-web .
docker run -p 5000:5000 convertly-web
```

## Configuration

The application uses environment variables for configuration. Key settings in `.env`:

```
# Flask
FLASK_DEBUG=False
SECRET_KEY=your_secret_key
FLASK_ENV=production

# Server
FLASK_PORT=5000
HOST=0.0.0.0

# File Upload
UPLOAD_FOLDER=uploads
MAX_CONTENT_LENGTH=33554432  # 32MB
FILE_EXPIRY_HOURS=24

# Payment
PAYPAL_CLIENT_ID=your_paypal_id
PAYPAL_CLIENT_SECRET=your_paypal_secret
PAYPAL_MODE=sandbox

# Claude API
CLAUDE_API_KEY=your_claude_api_key

# Freemium
FREE_CONVERSIONS_LIMIT=3
PAID_CONVERSIONS_AMOUNT=20

# Vouchers
VOUCHER_CODES=ITSD2026,BETA-ACCESS,CONVERTLY-VIP,TEAM50,FREE-TRIAL
VOUCHER_GRANT=50
```

## Project Structure

```
convertly-web/
├── app.py                    # Main Flask application
├── run.py                    # Application entry point
├── translation_service.py    # Translation functionality
├── test_features.py          # Feature tests
├── requirements.txt          # Python dependencies
├── Dockerfile                # Docker configuration
├── render.yaml               # Render deployment config
├── static/                   # Static assets (CSS, JS, fonts)
├── templates/                # HTML templates
│   ├── index.html           # Home page
│   ├── invoice.html         # Invoice page
│   └── share.html           # Share page
└── uploads/                  # User uploaded files (temporary)
```

## Usage

### Converting Files

1. Go to the home page (index.html)
2. Upload a file
3. Select target format
4. Click Convert
5. Download the converted file

### Free vs Premium

- **Free Tier**: Limited conversions per session/day
- **Premium/Paid**: Unlimited conversions with PayPal payment
- **Vouchers**: Use promotional codes for bonus conversions

## API Endpoints

- `POST /convert` - Convert uploaded file
- `POST /payment/create` - Create PayPal payment
- `GET /payment/success` - Handle PayPal success callback
- `GET /payment/cancel` - Handle PayPal cancellation
- `GET /share` - Generate shareable links for conversions
- `GET /sitemap.xml` - XML sitemap for SEO

## Deployment

### Render.com

The application is configured for deployment on Render using `render.yaml`:

```bash
git push origin main  # Triggers automatic deployment
```

### Environment Variables on Render

Set all `.env` variables in your Render project settings:
- CLAUDE_API_KEY
- PAYPAL_CLIENT_ID
- PAYPAL_CLIENT_SECRET
- VOUCHER_CODES
- etc.

## Security

- **CSP Headers**: Content Security Policy enabled
- **Session Security**: Secure, HttpOnly cookies
- **Secret Management**: API keys stored in environment variables
- **File Validation**: Uploaded files validated before processing
- **Rate Limiting**: Conversion rate limits to prevent abuse

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature`
3. Commit changes: `git commit -am 'Add new feature'`
4. Push to branch: `git push origin feature/your-feature`
5. Submit a Pull Request

## License

This project is licensed under the MIT License - see LICENSE file for details.

## Support

For issues, questions, or suggestions, please open a GitHub issue or contact the development team.

## Roadmap

- [ ] Additional file format support
- [ ] Batch file conversion
- [ ] Advanced compression options
- [ ] User accounts and history
- [ ] Mobile app
- [ ] API access for developers

---

**Built with ❤️ for easy file conversions**
