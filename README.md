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
`QuickSpreadsheet.call` accepts the following options:


     Option          | Format              
     --------------- |:--------------------
     `title`         | `Array<string>`      
     `header_row`    | `Array<string>`        
     `body_rows`     | `Array<Array<string>`
     `folder_path` (optional)   | `string`  

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


     List of Users
     Generated:  2016-02-03-19:30
     
     Id          | Name              
     ------------|--------------------
     1           | John Doe      
     2           | Jane Doe      
     3           | DHH           



That will be located in the `/tmp` folder in your app's root directory with the name of the spreadsheet and date time it was created:  `/tmp/ListofUsers_g<2016-02-03-16.09.xls`