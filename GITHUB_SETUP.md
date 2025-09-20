# GitHub Repository Setup Guide

This document provides instructions on how to push this project to GitHub.

## Prerequisites

1. **GitHub Account**: Ensure you have a GitHub account at [github.com](https://github.com)
2. **Git Installation**: Make sure Git is installed on your computer
3. **Repository Created**: Create a new repository on GitHub (e.g., "Code2Word" or "Report_generator")

## Steps to Push to GitHub

### 1. Initial Setup (First-time only)

```powershell
# Configure Git with your identity
git config --global user.name "Your Name"
git config --global user.email "your.email@example.com"
```

### 2. Initialize Git Repository (If not already done)

```powershell
# Navigate to your project directory
cd C:\Users\RAJESH\CODE\Python_code\code2word

# Initialize git repository
git init
```

### 3. Add Your GitHub Repository as Remote

```powershell
# Add the remote repository URL
git remote add origin https://github.com/yourusername/repository-name.git

# Verify the remote was added
git remote -v
```

### 4. Stage and Commit Files

```powershell
# Stage all files
git add .

# Commit with a descriptive message
git commit -m "Initial commit: Code2Word report generator with optimized performance"
```

### 5. Push to GitHub

```powershell
# Push to the main branch
git push -u origin main
```

### 6. Creating Releases (Optional)

After pushing your code, you can create a release on GitHub:

1. Go to your repository page on GitHub
2. Click on "Releases" in the right sidebar
3. Click "Create a new release"
4. Add a tag version (e.g., "v1.0.0")
5. Write a release title and description
6. Upload the `Code2word.exe` file from your `dist/` folder as a binary attachment
7. Click "Publish release"

## Troubleshooting

### Authentication Issues

If you encounter authentication issues, consider:
1. Using GitHub CLI to authenticate
2. Setting up SSH keys
3. Using a personal access token:
   ```powershell
   git remote set-url origin https://YOUR-USERNAME:YOUR-TOKEN@github.com/YOUR-USERNAME/YOUR-REPO.git
   ```

### Large File Issues

If you have trouble pushing the executable due to its size:
1. Add `dist/` to `.gitignore` if not already there
2. Use GitHub Releases to upload the executable separately
3. Consider using Git LFS for large files if needed

## Maintenance

Once your repository is set up, regular updates can be pushed with:

```powershell
git add .
git commit -m "Description of changes"
git push
```