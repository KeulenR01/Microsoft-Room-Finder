# Set-RoomFinder.ps1 — Step-by-step Room Finder Configuration (Exchange Online)

Configure Exchange Online **room mailboxes** so they appear correctly in **Outlook Room Finder** and can be filtered by location/capacity/equipment.  
This script guides an admin through a **structured, interactive, step-by-step** setup:

1. **Select room mailboxes**
2. **Review & complete Place metadata**
3. **Add rooms to a Room List (or create one)**  
4. **Show a summary** and optionally repeat for more rooms (without reconnecting)

---

## Why this script exists

Setting up Room Finder correctly often requires repetitive actions and careful attention to metadata consistency:

- Room mailboxes need **Place** properties (e.g., *City, Floor, Capacity, devices, accessibility, tags*).
- Rooms must be members of at least one **Room List** (a distribution group of type `RoomList`).  
  Without a Room List membership, a room may not show up in Room Finder.

This script provides a consistent, guided workflow so administrators can configure rooms reliably.

---

## What it does (high level)

### Step 1 — Inventory & select rooms
- Retrieves all room mailboxes (`Get-Mailbox -RecipientTypeDetails RoomMailbox`)
- Displays a numbered list
- Lets you select rooms using flexible input:
  - `all`
  - single number: `3`
  - comma-separated: `1,4,7`
  - range: `5-12`
  - combinations: `1,3-6,10`
- Validates each selected room for:
  - correct recipient type (`RoomMailbox`)
  - address list visibility (warns if hidden)
  - SMTP uniqueness checks (`Get-Recipient`)

<img width="1076" height="325" alt="image" src="https://github.com/user-attachments/assets/d4367b10-3d65-48ab-a06d-4ceeefb634c9" />


### Step 2 — Review & complete Place metadata
- Reads current Place data (`Get-Place`)
- Highlights missing values
- Prompts you to fill missing fields **field-by-field**
- Supports “suggested defaults” for some fields:
  - `CountryOrRegion`: type `/` to accept `NL`  
  - `AudioDeviceName`: type `/` to accept `Speaker`
  - `DisplayDeviceName`: type `/` to accept `TV Screen`
  - `VideoDeviceName`: type `/` to accept `Teams Camera`
  - `IsWheelChairAccessible`: type `/` to accept `false`
- **Tags are always prompted**, even if all other fields are filled:
  - `Teams Room`
  - `Whiteboard`
  - `HDMI`
- Applies changes using `Set-Place` and immediately verifies by re-reading `Get-Place`

> **Note:** Microsoft requires certain Place fields for Room Finder filtering to work well (commonly including *City, Floor, Capacity*). This script explicitly checks for missing fields.

<img width="1007" height="617" alt="image" src="https://github.com/user-attachments/assets/3b9c44bd-571a-47d1-8813-8bef60291a78" />


### Step 3 — Room Lists (add or create)
- Lists existing Room Lists (`Get-DistributionGroup` filtered by `RecipientTypeDetails -eq 'RoomList'`)
- For each selected room, you can:
  1. Add to an existing Room List
  2. Create a new Room List (with uniqueness checks)
  3. Skip
- Adds membership (`Add-DistributionGroupMember`)
- Verifies membership (`Get-DistributionGroupMember`)

### Summary
- Prints a per-room summary of:
  - remaining missing metadata
  - selected Room List
  - membership result
  - errors encountered

<img width="921" height="383" alt="image" src="https://github.com/user-attachments/assets/ccbfac9c-1889-442b-a1e2-bad706272b8c" />

---

## Requirements

- **Windows PowerShell 5.1**  
  (Script header: `#Requires -Version 5.1`)
- **ExchangeOnlineManagement module**
  ```powershell
  Install-Module ExchangeOnlineManagement

---

## License
This project is licensed under the [MIT License](LICENSE) — feel free to use, modify and distribute.

*Disclaimer*  
This repository contains generic, public best-practice scripts. It is not affiliated with any current or past employer or client. Created for educational and community purposes only.

