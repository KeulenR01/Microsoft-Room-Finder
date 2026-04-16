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

---

## Requirements

- **Windows PowerShell 5.1**  
  (Script header: `#Requires -Version 5.1`)
- **ExchangeOnlineManagement module**
  ```powershell
  Install-Module ExchangeOnlineManagement


# Configure the Microsoft Room Finder - the easy way!

Configuring rooms for Outlook Room Finder in Exchange Online is a repetitive and
error-prone process. Each room mailbox requires Place metadata (City, Floor,
Capacity, devices, accessibility, tags, etc.) to be set correctly before it
appears and filters properly in Room Finder. On top of that, every room must be
a member of at least one Room List (a distribution group of type RoomList) —
without this, the room is invisible to Room Finder entirely.

Doing this manually through the Exchange Admin Center or by running individual
PowerShell commands per room is time-consuming, easy to get wrong, and offers
no consistent overview of what has and has not been configured.

This script solves that by guiding an administrator interactively through the
full Room Finder setup in three structured steps:

Step 1 - Inventory and select room mailboxes
    Retrieves all room mailboxes, displays them in a numbered list, and
    lets the operator select one or more rooms using flexible input
    (all, single numbers, comma-separated, ranges, or combinations).
    Each selected room is validated for type, address-list visibility,
    and SMTP uniqueness before proceeding.
    
<img width="1076" height="325" alt="image" src="https://github.com/user-attachments/assets/0e1a2d78-d51e-4dd3-81b2-06ed88eae083" />
<img width="853" height="125" alt="image" src="https://github.com/user-attachments/assets/9777de10-9171-48f5-9ca6-f229f3501c4a" />


Step 2 - Verify and complete Place metadata
    Retrieves the current Place properties for each room via Get-Place
    and highlights missing fields. The operator is prompted to fill in
    missing values field by field. Fields with a suggested default can be
    accepted by typing / (slash). Tags are always offered as a numbered
    selection menu, even when all other fields are already filled.
    Changes are applied with Set-Place and verified immediately.
    
<img width="1007" height="617" alt="image" src="https://github.com/user-attachments/assets/fd565e9e-2c3d-4a30-8810-1cd6e62265c6" />


Step 3 - Link or create Room Lists
    Retrieves all existing Room Lists and lets the operator assign each
    room to an existing list or create a new one. Uniqueness checks are
    performed before creating a new Room List. Membership is verified
    after adding via Get-DistributionGroupMember.
    
<img width="921" height="383" alt="image" src="https://github.com/user-attachments/assets/9bcfc1d2-07fb-414e-a7d3-4ea4e2407add" />


At the end of each run a summary is printed per room. The operator can choose to run the script again for additional rooms without reconnecting.
