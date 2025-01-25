# React Fluent UI Virtual Multi Select Dropdown

## Introduction

This project is a MultiSelect Lookup component built using React and Fluent UI, designed specifically for Model-Driven Apps. It allows users to select multiple items from a lookup field, supporting only N:N (many-to-many) relationships between tables. The component enhances user experience by providing a seamless and efficient way to manage complex data relationships within the app.

## Configuration

To configure the project, you need to set the following properties:

### Property 1: `Related Entity Type`

- **Description**: Related entity logical name.
- **Default Value**: `account`

### Property 2: `Related Primary Columns`

- **Description**: Primary columns of the related table (include uniqueidentifier and name in below format).
- **Default Value**: `accoundid,name`

### Property 3: `Primary Entity Type`

- **Description**: Primary entity logical name. This is where control gets used and configured.
- **Default Value**: `contact`

### Property 4: `Relationship Name`

- **Description**: N:N relationship name which can be retrieved from primary or related entity.
- **Default Value**: `Account_Contact`

### Property 5: `Primary Entity ID`

- **Description**: Select primary entity id column from list.
- **Example**: Account (Text)

## Installation

Download latest version from repo and import manually or use pac CLI to import using terminal. For example:

```powershell
pac solution import --path c:\Users\Documents\Solution.zip
```

## Usage

After importing the solution, follow these steps to use the MultiSelect Lookup control:

1. **Choose a Text Field**: Pick any text field in your form where you want to use the MultiSelect Lookup control. You can use an existing text field or create a new one.

2. **Configure the Control**:

   - Open the form editor and select the text field.
   - In the field properties, go to the "Components" section.
   - Click on "+ Components" and select the imported MultiSelect Lookup control from the list.

3. **Set Properties**:

   - Configure the properties of the control as described in the Configuration section above.
   - Make sure to set the `Related Entity Type`, `Related Primary Columns`, `Primary Entity Type`, `Relationship Name`, and `Primary Entity ID` according to your requirements.

4. **Save and Publish**:
   - After setting the properties, save the form.
   - Publish the customizations to apply the changes.

That's it! Your MultiSelect Lookup control is now ready to use. It will allow users to select multiple items from the lookup field, making data management much easier and more efficient.

## How it works

https://www.youtube.com/watch?v=4wPPHakaq8I&t=4s
