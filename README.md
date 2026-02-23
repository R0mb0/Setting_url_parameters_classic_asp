# Setting url parameters Classic asp

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/768a54d9ee4f42f59f84480736c61c74)](https://app.codacy.com/gh/R0mb0/Setting_url_parameters_classic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Setting_url_parameters_classic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Setting_url_parameters_classic_asp)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

## `setting_parameters_to_url.class.asp`'s avaible functions

  - Add a url parameter to the list -> `Public Function add_parameter(ByVal name, ByVal parameter)`
  - Remove a parameter from the list by name -> `Public function remove_parameter_by_name(ByVal name)`
  - Remove a parameter from the list by the parameter value -> `Public function remove_parameter_by_parameter(ByVal parameter)`
  - Add the parameters to a url -> `Public Function set_parameters_to_url(ByVal url)`
  - Get actual page URL -> `Public Function get_current_url()`
  - Redirect a URL to new tab -> `Public Function redirect(ByVal url)`

## How to use

> From `Test.asp`

1. Initialize the class
   ```asp
    <%@LANGUAGE="VBSCRIPT"%>
    <!--#include file="setting_parameters_to_url.class.asp"-->
    <%
        'Initialize the class
        Dim my_parameters
        Set my_parameters = new setting_parameters_to_url
   ```

2. Use the class
    ```asp
        my_parameters.add_parameter "PFM", "Impressioni_di_settembre"
        my_parameters.add_parameter "Banco", "RIP"
        my_parameters.add_parameter "Locanda_delle_fate", "Forse_le_lucciole_non_si_amano_più"
    
        Response.write "<br><br>" & my_parameters.set_parameters_to_url(my_parameters.get_current_url + "/example.asp")
    
        my_parameters.remove_parameter_by_name "Banco"
    
        Response.write "<br><br>" & my_parameters.set_parameters_to_url(my_parameters.get_current_url + "/example.asp")
    
        my_parameters.remove_parameter_by_parameter "Impressioni_di_settembre"
    
        Response.write "<br><br>" & my_parameters.set_parameters_to_url(my_parameters.get_current_url + "/example.asp")
    %>
    ```

<a href="https://github.com/R0mb0/Not_made_by_AI">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="https://github.com/R0mb0/Not_made_by_AI/blob/main/Badge/SVG/NotMadeByAIDark.svg">
    <source media="(prefers-color-scheme: light)" srcset="https://github.com/R0mb0/Not_made_by_AI/blob/main/Badge/SVG/NotMadeByAILight.svg">
    <img alt="Not made by AI" src="https://github.com/R0mb0/Not_made_by_AI/blob/main/Badge/SVG/NotMadeByAIDefault.svg">
  </picture>
</a>
