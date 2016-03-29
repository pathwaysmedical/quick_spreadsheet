# quick_spreadsheet
Sensible defaults for fast excel spreadsheet creation.


## Installation

Add it to your Gemfile:

```ruby
gem 'quick_spreadsheet', '0.0.3', git: "git@github.com:pathwaysmedical/quick_spreadsheet.git"
```

Run the following command to install it:

```console
bundle install
```


## Usage
`QuickSpreadsheet.call` accepts two types of syntax.
Both are argument-order independent.

The first accepts the following options:


     Option          | Format              
     --------------- |:--------------------
     `title`         | `Array<string>`      
     `header_row`    | `Array<string>`        
     `body_rows`     | `Array<Array<string>`
     `filename` (optional)   | `string`  
     `folder_path` (optional)   | `string`  

The second accepts the following options:


     Option          | Format              |
     --------------- |:-------------------- |:--------------------
     `sheets`         | `Array<Array<`      |
     | `title`         | `Array<string>`      
     | `header_row`    | `Array<string>`        
     | `body_rows`     | `Array<Array<string>`
     `filename` (optional)   | `string`  |
     `folder_path` (optional)   | `string`  |

## Example

To create a spreadsheet of all users by `name` and `id` attributes, assuming a Rails ActiveRecord model of `User`:

```ruby
QuickSpreadsheet.call(
  title: "List of Users", 
  header_row: ["id","name"], 
  body_rows: User.all.map{|s| [s.id.to_s, s.name.to_s] } 
)
```

Will create a spreadsheet with this structure:

     List of Users (worksheet name)
     
     Id          | Name              
     ------------|--------------------
     1           | John Doe      
     2           | Jane Doe      
     3           | DHH           



The spreadsheet will be located in the `/tmp` folder in your app's root directory with the name of the spreadsheet and date time it was created:  `/tmp/ListofUsers_g<2016-02-03-16.09.xls`

To create a spreadsheet with a first worksheet being a list of all users by `name` and `id` attributes, assuming a Rails ActiveRecord model of `User`, and a second worksheet of those users' ids and emails attributes:

```ruby
QuickSpreadsheet.call(
  sheets: [
    {
      title: "List of Users", 
      header_row: ["id","name"], 
      body_rows: User.all.map{|s| [s.id.to_s, s.name.to_s] } 
    },
    {
      title: "User Emails", 
      header_row: ["id","email"], 
      body_rows: User.all.map{|s| [s.id.to_s, s.email.to_s] } 
    }
  ],
  folder_path: "~/User/documents/spreadsheets"
)
```

Will create a spreadsheet with this structure:

    "List of Users" (first worksheet name)
     
     Id          | Name              
     ------------|--------------------
     1           | John Doe      
     2           | Jane Doe      
     3           | DHH           


    "User Emails" (second worksheet name)
     
     Id          | Email              
     ------------|--------------------
     1           | john@doe.com      
     2           | jane@doe.com      
     3           | DHH@ror.org           



The spreadsheet will be located in the `/tmp` folder in your app's root directory with the name of the spreadsheet and date time it was created:  `~Users/documents/spreadsheets/ListofUsers_g<2016-02-03-16.09.xls`

