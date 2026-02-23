
# Select-ADObject

Interactive Active Directory object picker with GUI-based forest browsing and Tier 0 highlighting.

<img width="527" height="546" alt="image" src="https://github.com/user-attachments/assets/5c85d3a8-5613-4c81-ae12-2c29e853337f" />




> **Note**  
> The script parameters are specifically tailored for my AS2Go PowerShell module, and the icons are optimized for PowerShell 7.

---

## Overview

`Select-ADObject` is a PowerShell function that provides a graphical interface for browsing and selecting Active Directory objects within a forest.

It enables administrators to:

- Browse multiple domains in a forest
- Select OUs, users, groups, computers, or domains
- Restrict scope to the local domain
- Highlight Tier 0 objects
- Control selection scope via switches
- Return the distinguished name (DN) of the selected object

The function is designed for administrative tooling, lab environments, and security-focused workflows.

---

## Requirements

- Windows environment (GUI required)
- PowerShell 5.1+ or PowerShell 7 (Windows)
- Active Directory module
- Access to query AD forest structure
- System.Windows.Forms
- System.DirectoryServices

---

## Installation

Copy the function into:

- A PowerShell module
- Your profile
- A script library

```powershell
. .\Select-ADObject.ps1
```

---

## Usage

### Example #1

``` powershell
Select-ADObject
```

### Example #2

```powershell
Select-ADObject -MarkTier0 -IncludeUsers -Title "Select Break Glass Account"
```

---

## Parameters

<details>
<summary><strong>Selection Scope</strong></summary>

<br>

| Parameter | Type | Default | Description |
|------------|--------|----------|-------------|
| `Title` | String | `"Select AD Object"` | Title displayed in the selection dialog window. |
| `LocalDomainOnly` | Switch | `False` | Displays only the current domain and its objects. |
| `AllowDomainSwitch` | Switch | `False` | Allows domain objects themselves to be selected. |
| `DomainSelectionOnly` | Switch | `False` | Restricts selection strictly to domain objects. |

</details>

---

<details>
<summary><strong>Object Types</strong></summary>

<br>

| Parameter | Type | Default | Description |
|------------|--------|----------|-------------|
| `IncludeUsers` | Switch | `False` | Includes user objects in the tree view. |
| `IncludeGroups` | Switch | `False` | Includes group objects in the tree view. |
| `IncludeComputers` | Switch | `False` | Includes computer objects in the tree view. |

</details>

---

<details>
<summary><strong>Security / Tier 0 Handling</strong></summary>

<br>

| Parameter | Type | Default | Description |
|------------|--------|----------|-------------|
| `MarkTier0` | Switch | `False` | Highlights detected Tier 0 objects with a red indicator. |
| `Tier0Regex` | String | Default Tier 0 pattern | Regular expression used to identify Tier 0 objects. |

**Default Tier 0 Pattern:**

```regex
(?i)(^CN=Administrator,|^OU=Domain Controllers,|^OU=Tier\s*0|^OU=T0)
```

</details>

---

## Output

| Type | Description |
|------|------------|
| `System.String` | Returns the Distinguished Name (DN) of the selected object. Returns `$null` if cancelled. |

---

## Security Considerations

- Requires directory read permissions
- Tier 0 detection is regex-based (customizable)
- Does not modify AD
- No credentials stored
- Intended for administrative usage

---

## Limitations

- Requires GUI environment
- Not compatible with headless execution
- Performance depends on forest size

---

## License

Internal / Custom usage unless otherwise specified.
