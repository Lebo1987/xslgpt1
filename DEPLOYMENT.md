# Deployment Guide for XSLGPT1 Excel Add-in

This guide will help you deploy your Excel Add-in to Render and set up Git version control.

## 🚀 Quick Deployment Steps

### 1. Create GitHub Repository

1. Go to [GitHub](https://github.com) and sign in
2. Click the "+" icon in the top right and select "New repository"
3. Name it `xslgpt1-excel-addin`
4. Make it **Public** (required for free Render deployment)
5. **Don't** initialize with README (we already have one)
6. Click "Create repository"

### 2. Push to GitHub

Run these commands in your project directory:

```bash
# Add GitHub as remote origin (replace YOUR_USERNAME with your GitHub username)
git remote add origin https://github.com/YOUR_USERNAME/xslgpt1-excel-addin.git

# Push to GitHub
git branch -M main
git push -u origin main
```

### 3. Deploy to Render

1. Go to [Render](https://render.com) and sign up/sign in
2. Click "New +" and select "Web Service"
3. Connect your GitHub account if not already connected
4. Select your `xslgpt1-excel-addin` repository
5. Configure the service:
   - **Name**: `xslgpt1-excel-addin`
   - **Environment**: `Node`
   - **Build Command**: `npm install && npm run build`
   - **Start Command**: `npm start`
   - **Plan**: `Free`
6. Click "Create Web Service"

### 4. Update Manifest for Production

After deployment, Render will give you a URL like:
`https://xslgpt1-excel-addin.onrender.com`

Update the `manifest-production.xml` file with your actual Render URL:

```xml
<!-- Replace all instances of: -->
https://xslgpt1-excel-addin.onrender.com

<!-- With your actual Render URL -->
```

### 5. Test the Deployment

1. Visit your Render URL to verify the server is running
2. Test the API endpoint: `https://your-app.onrender.com/api/generate`
3. Use the production manifest in Excel

## 📁 Project Structure After Deployment

```
XSLGPT1/
├── src/                    # Source code
├── dist/                   # Built files (generated)
├── assets/                 # Icons
├── server.js              # Backend server
├── package.json           # Dependencies
├── webpack.config.js      # Build config
├── render.yaml            # Render configuration
├── manifest-local.xml     # Local development manifest
├── manifest-production.xml # Production manifest
├── .gitignore            # Git ignore rules
└── README.md             # Documentation
```

## 🔧 Configuration Files

### render.yaml
```yaml
services:
  - type: web
    name: xslgpt1-excel-addin
    env: node
    plan: free
    buildCommand: npm install && npm run build
    startCommand: npm start
    envVars:
      - key: NODE_ENV
        value: production
      - key: PORT
        value: 10000
```

### Environment Variables
Create these in Render dashboard:
- `NODE_ENV`: `production`
- `PORT`: `10000` (or let Render assign automatically)

## 🌐 URLs After Deployment

- **Web Service**: `https://your-app-name.onrender.com`
- **API Endpoint**: `https://your-app-name.onrender.com/api/generate`
- **Taskpane**: `https://your-app-name.onrender.com/taskpane.html`

## 📝 Using the Deployed Add-in

### For Development (Local)
1. Use `manifest-local.xml`
2. Server runs on `localhost:8080`

### For Production (Deployed)
1. Use `manifest-production.xml`
2. Server runs on Render URL

## 🔄 Continuous Deployment

Render will automatically:
- Deploy when you push to `main` branch
- Rebuild when `package.json` changes
- Restart on server crashes

## 🐛 Troubleshooting

### Common Issues

1. **Build Fails**
   - Check `package.json` has all dependencies
   - Verify `npm run build` works locally

2. **Add-in Not Loading**
   - Ensure HTTPS URLs in manifest
   - Check CORS settings in server.js

3. **API Errors**
   - Verify server is running on Render
   - Check environment variables

4. **Git Issues**
   - Ensure repository is public (for free Render)
   - Check remote origin is correct

### Debug Commands

```bash
# Check Git status
git status

# Check remote origin
git remote -v

# Test build locally
npm run build

# Test server locally
npm start
```

## 📊 Monitoring

- **Render Dashboard**: Monitor logs and performance
- **GitHub**: Track code changes and issues
- **Excel**: Test add-in functionality

## 🔐 Security Notes

- Keep API keys in environment variables
- Use HTTPS for all production URLs
- Don't commit `.env` files to Git

## 🚀 Next Steps

1. **Custom Domain**: Add custom domain in Render
2. **SSL Certificate**: Automatically provided by Render
3. **Scaling**: Upgrade to paid plan for more resources
4. **Monitoring**: Add logging and analytics

## 📞 Support

- **Render Docs**: https://render.com/docs
- **Office Add-ins**: https://docs.microsoft.com/en-us/office/dev/add-ins/
- **GitHub Issues**: Create issues in your repository

---

**Happy Deploying! 🎉** 