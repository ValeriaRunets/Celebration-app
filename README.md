# Celebration App - PowerApps Canvas App

A PowerApps Canvas App for managing employee celebration events (e.g., new joiner welcomes). Managers can review pre-drafted emails, set communication dates, and submit celebrations for review.

## Screen: Celebration Detail

The main screen allows managers to:
- View event summary (celebration type, employee name, status)
- Edit a pre-drafted manager welcome email
- Set a communication trigger date
- Select a communication sender
- Optionally draft a personal email
- Submit for review, save as draft, or decline

## Project Structure

```
├── CanvasManifest.json                 # App manifest (name, layout, screen order)
├── Properties.json                     # App properties and configuration
├── src/
│   ├── App.fx.yaml                     # App-level formulas, theme, and sample data
│   └── CelebrationDetailScreen.fx.yaml # Main screen with all controls
└── src/Assets/Images/                  # Placeholder for image assets
```

## How to Import into Power Apps

### Option 1: Using Power Platform CLI (pac)
1. Install the [Power Platform CLI](https://learn.microsoft.com/en-us/power-platform/developer/cli/introduction)
2. Pack the source files into an .msapp:
   ```bash
   pac canvas pack --sources src --msapp CelebrationApp.msapp
   ```
3. Open [Power Apps Studio](https://make.powerapps.com)
4. Go to **Apps** > **Import canvas app** > upload the `.msapp` file

### Option 2: Manual Recreation
1. Open [Power Apps Studio](https://make.powerapps.com)
2. Create a new **Canvas App** (Tablet layout, 1366x768)
3. Recreate the controls following the YAML definitions in `src/CelebrationDetailScreen.fx.yaml`
4. Copy the `App.OnStart` formula from `src/App.fx.yaml`

## Color Palette

| Element | Color | Hex |
|---------|-------|-----|
| Header Background | Dark Maroon | `#4A1942` |
| Info Banner Background | Light Orange | `#FFF3E0` |
| Info Banner Text | Dark Orange | `#E65100` |
| Joiner Badge | Teal | `#00897B` |
| Status "Needs Input" | Orange-Red | `#D84315` |
| Decline Button | Red | `#D32F2F` |
| Submit Button | Dark Gray | `#2D2D2D` |
| Body Text | Dark Gray | `#333333` |
| Label Text | Medium Gray | `#666666` |

## Data Model

The app uses a local collection `CelebrationEvents` with fields:
- `ID` - Event identifier
- `CelebrationType` - Type of celebration (e.g., "Joiner")
- `EmployeeFirstName` / `EmployeeLastName` - Employee details
- `Status` - Current event status
- `ManagerName` - Responsible manager
- `TriggerDate` - Communication trigger date
- `ManagerDraftEmail` - Pre-drafted email text
- `CommunicationSender` - Selected sender
- `MyEmail` - Optional personal email from manager
