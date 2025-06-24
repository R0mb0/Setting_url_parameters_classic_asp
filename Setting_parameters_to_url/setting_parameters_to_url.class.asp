<%
    Class setting_parameters_to_url

        '==========================
        'This is my first idea for something like this 
        '==========================

        'Private fields 

        'name array
        Dim names 
        'parameters array
        Dim parameters 
        'field for arrays status 
        Dim has_been_initialized 

        'Initialization and destruction'
        Sub class_initialize()
            names =  Array()
            parameters = Array()
            has_been_initialized = False
        End Sub 

        Sub class_terminate()
            names =  nothing 
            parameters = nothing 
            has_been_initialized = nothing 
        End Sub 

        'Function to add paramters to array 
        Public Function add_parameter(ByVal name, ByVal parameter)
            If Not(has_been_initialized) Then 
                'If the arrays has not been initialized 
                Redim Preserve names(0)
                Redim Preserve parameters(0)
                has_been_initialized = True 
                names(0) = name 
                parameters(0) = parameter 
            Else 
                'If the arrays has been initialized 
                Redim Preserve names(UBound(names) + 1)
                names(UBound(names)) = name 
                Redim Preserve parameters(UBound(parameters) + 1)
                parameters(UBound(parameters)) = parameter

            End If 
        End Function 

        'Funtion to remove by name
        Public function remove_parameter_by_name(ByVal name)
            Dim temp 
            Dim index 
            index = 0 
            Dim array_index 
            array_index = 0
            Dim temp_array
            temp_array = Array()
            Dim targhet_index
            Dim error
            error = True  

            For Each temp In names 
                If Not(temp = name) Then 
                    Redim Preserve temp_array(array_index)
                    temp_array(array_index) = temp 
                    array_index = array_index + 1
                Else 
                    targhet_index = index 
                    error = False 
                End If 
                index = index + 1
            Next 

            If error Then 
                 Call Err.Raise(vbObjectError + 10, "setting_parameters_to_url.class - remove_parameter_by_name", "The name is not present")
            End If 
            
            'Update array
            names = temp_array
            Dim targhet_parameter
            targhet_parameter = parameters(targhet_index)
            array_index = 0

            For Each temp In parameters
                If Not(temp = targhet_parameter) Then 
                    Redim Preserve temp_array(array_index)
                    temp_array(array_index) = temp 
                    array_index = array_index + 1
                End If 
            Next 

            'Update array
            parameters = temp_array
        End Function 


        'Funtion to remove by parameter
        Public function remove_parameter_by_parameter(ByVal parameter)
            Dim temp 
            Dim index 
            index = 0 
            Dim array_index 
            array_index = 0
            Dim temp_array
            temp_array = Array()
            Dim targhet_index
            Dim error 
            error = True 
            
            For Each temp In parameters 
                If Not(temp = parameter) Then 
                    Redim Preserve temp_array(array_index)
                    temp_array(array_index) = temp 
                    array_index = array_index + 1
                Else 
                    targhet_index = index 
                    error = False 
                End If 
                index = index + 1
            Next 

            If error Then 
                Call Err.Raise(vbObjectError + 10, "setting_parameters_to_url.class - remove_parameter_by_name", "The parameter is not present")
            End If 
            
            'Update array
            parameters = temp_array
            Dim targhet_parameter
            targhet_parameter = names(targhet_index)
            array_index = 0

            For Each temp In names
                If Not(temp = targhet_parameter) Then 
                    Redim Preserve temp_array(array_index)
                    temp_array(array_index) = temp 
                    array_index = array_index + 1
                End If 
            Next 

            'Update array
            names = temp_array
        End Function

        'Funtion to generate url with paramters 
        Public Function set_parameters_to_url(ByVal url)

            Dim my_url 
            my_url = url 
            my_url = my_url + "?"
            Dim index 
            Dim is_first
            is_first = True 

            For index = 0 To UBound(names)
                If is_first Then 
                    my_url = my_url + names(index) + "=" + parameters(index)
                    is_first = False 
                Else
                    my_url = my_url + "&" + names(index) + "=" + parameters(index)
                End If 
            Next

            set_parameters_to_url = my_url
        End Function 

         'Function to get current url'
        Public Function get_current_url()
            Dim protocol
            Dim domainName
            Dim fileName
            Dim queryString
            Dim url

            protocol = "http" 
            If lcase(request.ServerVariables("HTTPS"))<> "off" Then 
                protocol = "https" 
            End If

            domainName = Request.ServerVariables("SERVER_NAME") 
            fileName = Request.ServerVariables("SCRIPT_NAME") 
            queryString = Request.ServerVariables("QUERY_STRING")

            url = protocol & "://" & domainName & fileName
            If Len(queryString)<>0 Then
                url = url & "?" & queryString
            End If

            get_current_url = url 
        End Function

        'Function to redirect to another page'
        Public Function redirect(ByVal url)
            %>
                <SCRIPT language='javascript'>window.open('<%=url%>');</SCRIPT>
            <%
        End Function


        'For debug purpose 
        Public Function get_names
            For Each temp In names 
                Response.write temp & "<br>"
            Next 
        End Function 

        Public Function get_parameters
           For Each temp In parameters 
                Response.write temp & "<br>"
            Next 
        End Function 


    End Class
%> 