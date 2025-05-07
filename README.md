
# 🔁 CRM Access-Based Sales Team Tracking Plugin

This Dynamics 365 Dataverse plugin automatically manages Sales Team tracking based on user access grants or revocations to business records. It ensures real-time reflection of which users are actively participating in sales efforts, based on dynamic record-sharing actions.

---

## ✅ Features

- Automatically **creates tracking entries** when users are granted access
- **Marks tracking entries as inactive** when access is revoked
- Supports multiple business record types (e.g., applications, leads, opportunities)
- Links users to contextual business hierarchies (e.g., project, package)
- Captures user metadata (e.g., role, line of business)
- Supports both individual users and teams as access recipients

---

## 📈 Business Logic Flow

### 1. **Initialization**
- Initializes CRM services and plugin context
- Detects whether the trigger is an access grant or revoke

### 2. **On Access Grant**
- Retrieves the record and user who received access
- Checks if a Sales Team tracking entry already exists for this user and record
- If not:
  - Creates a new entry
  - Populates it with start date and metadata
  - Includes project or hierarchy-level context if available

### 3. **On Access Revoke**
- Retrieves the latest active Sales Team tracking record for this user and record
- Sets an end date to indicate the user is no longer active in the record

---

## 🔒 Security Requirements

- Plugin should be registered for:
  - `GrantAccess` (PostOperation)
  - `RevokeAccess` (PostOperation)
- Requires access to:
  - Tracking entity (Create, Update)
  - User metadata entity (Read)
  - Target record types (Read)

---

## 🧠 Technical Design Notes

- Uses conditional queries to avoid duplicate tracking entries
- Timestamps access lifecycle using start and end date fields
- Works generically across entities with business relevance
- Uses structured metadata (like role and responsibility) to enrich tracking entries
- Accessed metadata includes:
  - Role
  - Line of Business
  - Linked business objects (e.g., project/package)

---

## 📁 Suggested Registration Settings

| Plugin Action     | Message      | Execution Mode | Trigger Scope |
|-------------------|--------------|----------------|---------------|
| Access Granted    | GrantAccess  | PostOperation  | All Entities  |
| Access Revoked    | RevokeAccess | PostOperation  | All Entities  |

---

## 🧪 Example Use Case

When access is granted to a sales-related record:

1. A tracking record is created noting the user’s involvement and start date.
2. The entry includes associated business metadata and linked references.
3. If access is later revoked:
   - The plugin updates that tracking entry to mark the user as inactive.

---


