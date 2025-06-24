<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="setting_parameters_to_url.class.asp"-->
<%
    'Initialize the class
    Dim my_parameters
    Set my_parameters = new setting_parameters_to_url

    my_parameters.add_parameter "PFM", "Impressioni_di_settembre"
    my_parameters.add_parameter "Banco", "RIP"
    my_parameters.add_parameter "Locanda_delle_fate", "Forse_le_lucciole_non_si_amano_più"

    Response.write "<br><br>" & my_parameters.set_parameters_to_url(my_parameters.get_current_url + "/example.asp")

    my_parameters.remove_parameter_by_name "Banco"

    Response.write "<br><br>" & my_parameters.set_parameters_to_url(my_parameters.get_current_url + "/example.asp")

    my_parameters.remove_parameter_by_parameter "Impressioni_di_settembre"

    Response.write "<br><br>" & my_parameters.set_parameters_to_url(my_parameters.get_current_url + "/example.asp")
%>