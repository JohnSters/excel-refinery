# ExcelRefinery - Technical Architecture Documentation
**Version:** 1.0  
**Date:** January 2025  
**Project:** ExcelRefinery  

## Table of Contents
1. [Project Overview](#project-overview)
2. [Technical Stack](#technical-stack)
3. [Project Structure](#project-structure)
4. [Dependencies & Libraries](#dependencies--libraries)
5. [Configuration](#configuration)
6. [Database Architecture](#database-architecture)
7. [Authentication & Security](#authentication--security)
8. [Development Environment](#development-environment)
9. [Containerization](#containerization)
10. [Build & Deployment](#build--deployment)
11. [Key Features & Capabilities](#key-features--capabilities)

## Project Overview

ExcelRefinery is an ASP.NET Core 8.0 web application designed for Excel and CSV data processing and manipulation. The application follows the Model-View-Controller (MVC) architectural pattern with integrated user authentication and database storage capabilities.

### Core Purpose
- Excel file processing and data refinement
- CSV data handling and transformation
- User authentication and session management
- Web-based interface for data operations

## Technical Stack

### Core Framework
- **Framework:** ASP.NET Core 8.0
- **Runtime:** .NET 8.0
- **Language:** C# with nullable reference types enabled
- **Architecture Pattern:** MVC (Model-View-Controller)

### Database
- **ORM:** Entity Framework Core 8.0.17
- **Database Provider:** SQL Server
- **Local Development:** SQL Server LocalDB
- **Migration Support:** Code-First with EF Migrations

### Frontend
- **UI Framework:** Bootstrap (included via libman)
- **JavaScript:** jQuery with validation libraries
- **CSS:** Custom site styling with Bootstrap theming
- **View Engine:** Razor Views (.cshtml)

### Authentication
- **Identity System:** ASP.NET Core Identity
- **User Management:** IdentityUser with EntityFramework stores
- **Account Confirmation:** Email confirmation required

## Project Structure

```
ExcelRefinery/
├── Controllers/           # MVC Controllers
│   └── HomeController.cs
├── Data/                 # Database Context & Migrations
│   ├── ApplicationDbContext.cs
│   └── Migrations/
├── Models/               # Data Models & ViewModels
│   └── ErrorViewModel.cs
├── Views/                # Razor Views
│   ├── Home/
│   └── Shared/
├── Areas/                # Feature Areas
│   └── Identity/         # Identity UI Pages
├── wwwroot/              # Static Files
│   ├── css/
│   ├── js/
│   └── lib/
├── Properties/           # Project Configuration
└── Docs/                 # Documentation
    └── Technical/
```

## Dependencies & Libraries

### Core ASP.NET Core Packages
```xml
Microsoft.AspNetCore.Diagnostics.EntityFrameworkCore (8.0.17)
Microsoft.AspNetCore.Identity.EntityFrameworkCore (8.0.17)
Microsoft.AspNetCore.Identity.UI (8.0.17)
Microsoft.EntityFrameworkCore.SqlServer (8.0.17)
Microsoft.EntityFrameworkCore.Tools (8.0.17)
```

### Excel & Data Processing Libraries
```xml
ClosedXML (0.105.0)                    # Excel file creation and manipulation
ExcelDataReader (3.7.0)               # Excel file reading capabilities
CsvHelper (33.1.0)                    # CSV file processing
System.Text.Encoding.CodePages (9.0.7) # Text encoding support
```

### Development & Deployment
```xml
Microsoft.VisualStudio.Azure.Containers.Tools.Targets (1.22.1)
```

## Configuration

### Application Settings
- **Environment:** Development/Production configuration
- **Connection Strings:** SQL Server LocalDB for development
- **Logging:** Information level with ASP.NET Core warnings
- **Host Configuration:** All hosts allowed

### Development Ports
- **HTTP:** localhost:5169
- **HTTPS:** localhost:7059
- **IIS Express:** localhost:36373 (SSL: 44303)
- **Docker:** 8080 (HTTP), 8081 (HTTPS)

### User Secrets
- **Secret ID:** `aspnet-ExcelRefinery-780df508-08e0-4d9f-a06b-e20ec499c8fc`
- Used for storing sensitive configuration during development

## Database Architecture

### Database Context
- **Class:** `ApplicationDbContext`
- **Base:** `IdentityDbContext`
- **Provider:** SQL Server via Entity Framework Core

### Schema
The application uses ASP.NET Core Identity schema with the following tables:
- **AspNetUsers** - User accounts and profile information
- **AspNetRoles** - User roles and permissions
- **AspNetUserRoles** - User-role relationships
- **AspNetUserClaims** - User-specific claims
- **AspNetRoleClaims** - Role-based claims
- **AspNetUserLogins** - External login providers
- **AspNetUserTokens** - Authentication tokens

### Migration Management
- Code-First approach with EF Migrations
- Initial migration creates complete Identity schema
- Developer page exception filter enabled for development

## Authentication & Security

### Identity Configuration
- **User Type:** `IdentityUser` (default ASP.NET Core Identity)
- **Account Confirmation:** Required via email
- **Password Policy:** Default ASP.NET Core Identity rules
- **Two-Factor Authentication:** Supported via Identity framework

### Security Features
- HTTPS redirection enforced
- HSTS (HTTP Strict Transport Security) in production
- Anti-forgery token validation
- Secure cookie configuration

### Authorization
- Route-based authorization via `[Authorize]` attributes
- Role-based access control capability
- Claims-based authorization support

## Development Environment

### Required Tools
- **.NET 8.0 SDK** - Framework and runtime
- **Visual Studio 2022** or **VS Code** - IDE support
- **SQL Server LocalDB** - Local database development
- **Docker Desktop** - Container development (optional)

### Development Profiles
1. **Project (HTTP)** - Direct Kestrel hosting
2. **Project (HTTPS)** - Secure Kestrel hosting  
3. **IIS Express** - IIS development server
4. **Container** - Docker containerized development

### Environment Variables
- `ASPNETCORE_ENVIRONMENT=Development`
- `ASPNETCORE_HTTP_PORTS=8080` (Docker)
- `ASPNETCORE_HTTPS_PORTS=8081` (Docker)

## Containerization

### Docker Configuration
- **Base Image:** `mcr.microsoft.com/dotnet/aspnet:8.0`
- **Build Image:** `mcr.microsoft.com/dotnet/sdk:8.0`
- **Target OS:** Linux containers
- **Multi-stage Build:** Optimized for production deployment

### Container Ports
- **HTTP:** 8080
- **HTTPS:** 8081
- **SSL Support:** Enabled in container profile

### Build Process
1. **Restore** - NuGet package restoration
2. **Build** - Release configuration compilation
3. **Publish** - Application publishing for deployment
4. **Runtime** - Final container with published application

## Build & Deployment

### Build Configuration
- **Target Framework:** net8.0
- **Nullable Reference Types:** Enabled
- **Implicit Usings:** Enabled
- **Default Configuration:** Release for production

### Static File Management
- CSS bundling and minification
- JavaScript optimization
- Bootstrap and jQuery via libman
- Scoped CSS support for Razor components

### Production Considerations
- **Error Handling:** Custom error pages in production
- **HSTS:** 30-day max age (configurable)
- **Static File Caching:** Optimized for performance
- **Database Migrations:** Run via migrations endpoint

## Key Features & Capabilities

### Excel Processing
- **Read Operations:** Excel file parsing and data extraction
- **Write Operations:** Excel file generation and modification
- **Format Support:** XLSX, XLS file formats
- **Advanced Features:** Cell formatting, formulas, charts (via ClosedXML)

### CSV Processing
- **Parsing:** Flexible CSV reading with customizable delimiters
- **Generation:** CSV file creation and export
- **Data Mapping:** Object-to-CSV and CSV-to-object mapping
- **Encoding Support:** Various text encodings including code pages

### Web Interface
- **Responsive Design:** Bootstrap-based UI framework
- **User Authentication:** Complete registration and login system
- **File Upload:** Support for file upload and processing
- **Data Visualization:** Potential for charts and data display

### Extensibility Points
- **Custom Controllers:** Easy addition of new functionality
- **Custom Models:** Domain-specific data models
- **Service Layer:** Dependency injection ready
- **Custom Views:** Flexible UI customization

---

**Document Metadata:**
- **Author:** System Generated
- **Last Updated:** January 2025
- **Version Control:** Track changes in version control system
- **Related Documents:** User Guide, API Documentation, Deployment Guide 