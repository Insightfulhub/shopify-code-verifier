# Shopify Code Verifier

This Node.js app verifies unique product codes from an Excel file.

## 🚀 Deployment (Render.com)

1. Go to [https://render.com](https://render.com)
2. Create a **New Web Service**
3. Connect your GitHub repo (upload these files)
4. Set the **Start Command**: `node server.js`
5. Upload your `codes.xlsx` file (included)
6. Deploy!

### ✅ API Example

**GET Request:**
```
https://yourapp.onrender.com/verify?code=ABC123&name=John&mobile=9876543210
```

**Response:**
```json
{ "success": true, "message": "✅ Code verified successfully!" }
```

### 🛍 Shopify Integration
Use this API URL in your Shopify custom page JavaScript.
