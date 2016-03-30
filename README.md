# quick_spreadsheet
Sensible defaults for fast Excel spreadsheet creation.
Designed for use with Ruby on Rails applications.

## Installation

Assuming you have Bundler installed, add the gem to your Gemfile:

```ruby
gem 'quick_spreadsheet', '0.1.0', git: "git@github.com:pathwaysmedical/quick_spreadsheet.git"
```

Run the following command to install it:

```console
bundle install
```


## Usage

### Call
`QuickSpreadsheet.call` accepts two syntax variants. Both are argument-order independent.

The first accepts the following options:


     Option          | Format              
     --------------- |:--------------------
     `title`         | `Array<string>`      
     `header_row`    | `Array<string>`        
     `body_rows`     | `Array<Array<string>`
     `filename` (optional)   | `string`  
     `folder_path` (optional)   | `string`  

The second accepts the following options:


     Option          | Contains              | Format              
     --------------- |:-------------------- |:--------------------
     `sheets`         |               | `Array<`      
     | `title`         | `Array<string>`      
     | `header_row`    | `Array<string>`        
     | `body_rows`     | `Array<Array<string>`
     `filename` (optional)   |           | `string`  
     `folder_path` (optional)   |           | `string`  



### Args
`QuickSpreadsheet.args` returns the acceptable arguments, as per the tables above.

## Examples

### Single worksheet example

Here is an example using the simplest `QuickSpreadsheet.call` possible. To create a spreadsheet of all users by `id` and `name` attributes (assuming a Rails ActiveRecord model of `User`):

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



The spreadsheet will be located in the `/tmp` folder in your app's root directory with the name of the spreadsheet and date time it was created:  `/tmp/ListofUsers_g<2016-02-03-16.09.xls`.

### Mulitiple worksheets example

Here is an example using the nested `sheets` syntax, and also the optional `folder_path`. To create a spreadsheet with a first worksheet being a list of all users by `id` and `name` attributes (assuming a Rails ActiveRecord model of `User`), and a second worksheet of those users' `id` and `email` attributes:

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



The spreadsheet will be located in the folder specified, with the name of the spreadsheet and date time it was created:  `~/Users/documents/spreadsheets/ListofUsers_g<2016-02-03-16.09.xls`.
In this example, the file is named based upon the `title` of the first worksheet. To give it a different name, pass in the `filename` option, just as with the `folder_path` passed here.

