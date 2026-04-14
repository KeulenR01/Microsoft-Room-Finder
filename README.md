# Microsoft-Room-Finder

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

      Step 1 — Inventory and select room mailboxes
               Retrieves all room mailboxes, displays them in a numbered list, and
               lets the operator select one or more rooms using flexible input
               (all, single numbers, comma-separated, ranges, or combinations).
               Each selected room is validated for type, address-list visibility,
               and SMTP uniqueness before proceeding.

      Step 2 — Verify and complete Place metadata
               Retrieves the current Place properties for each room via Get-Place
               and highlights missing fields. The operator is prompted to fill in
               missing values field by field. Fields with a suggested default can be
               accepted by typing / (slash). Tags are always offered as a numbered
               selection menu, even when all other fields are already filled.
               Changes are applied with Set-Place and verified immediately.

      Step 3 — Link or create Room Lists
               Retrieves all existing Room Lists and lets the operator assign each
               room to an existing list or create a new one. Uniqueness checks are
               performed before creating a new Room List. Membership is verified
               after adding via Get-DistributionGroupMember.

    At the end of each run a summary is printed per room. The operator can choose
    to run the script again for additional rooms without reconnecting.
