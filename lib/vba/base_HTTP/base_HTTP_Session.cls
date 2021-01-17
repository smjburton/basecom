Option Explicit

' A Requests session.
' Provides cookie persistence, connection-pooling, and configuration.
' The Session object allows you to persist certain parameters across requests. It also
' persists cookies across all requests made from the Session instance. If you’re making several
' requests to the same host, the underlying TCP connection will be reused, which can result in
' a significant performance increase.

' See: http://docs.python-requests.org/en/latest/user/advanced/

	' Sending Multiple Requests With the Same Client¶
	' Zend\Http\Client was also designed specifically to handle several consecutive requests with the same object. This is useful in cases where a script requires data to be fetched from several places, or when accessing a specific HTTP resource requires logging in and obtaining a session cookie, for example.


	' If your application requires one authentication request per user, and consecutive requests might be performed in more than one script in your application, it might be a good idea to store the Cookies object in the user’s session. This way, you will only need to authenticate the user once every session.

	' Performing consecutive requests with one client

	' // First, instantiate the client
	' $client = new Zend\Http\Client('http://www.example.com/fetchdata.php', array(
	'     'keepalive' => true
	' ));

	' // Do we have the cookies stored in our session?
	' if (isset($_SESSION['cookiejar']) &&
	' 	$_SESSION['cookiejar'] instanceof Zend\Http\Cookies) {

    	' 	$cookieJar = $_SESSION['cookiejar'];
	' } else {
    	' // If we don't, authenticate and store cookies
    	' $client->setUri('http://www.example.com/login.php');
    	' $client->setParameterPost(array(
        'user' => 'shahar',
        'pass' => 'somesecret'
    	'  ));
    	' $response = $client->setMethod('POST')->send();
    	' $cookieJar = Zend\Http\Cookies::fromResponse($response);

    	' // Now, clear parameters and set the URI to the original one
    	' // (note that the cookies that were set by the server are now
    	' // stored in the jar)
    	' $client->resetParameters();
    	' $client->setUri('http://www.example.com/fetchdata.php');
    	' }

    	' // Add the cookies to the new request
    	' $client->setCookies($cookieJar->getMatchingCookies($client->getUri()));
    	' $response = $client->setMethod('GET')->send();

    	' // Store cookies in session, for next page
    	' $_SESSION['cookiejar'] = $cookieJar;


	' When performing several requests to the same host, it is highly recommended to enable the ‘keepalive’ configuration flag. This way, if the server supports keep-alive connections, the connection to the server will only be closed once all requests are done and the Client object is destroyed. This prevents the overhead of opening and closing TCP connections to the server.

	' When you perform several requests with the same client, but want to make sure all the request-specific parameters are cleared, you should use the resetParameters() method. This ensures that GET and POST parameters, request body and headers are reset and are not reused in the next request.

	' Note
	' Resetting parameters
	' Note that cookies are not reset by default when the resetParameters() method is used. To clean all cookies as well, use resetParameters(true), or call clearCookies() after calling resetParameters().
	' Another feature designed specifically for consecutive requests is the Zend\Http\Cookies object. This “Cookie Jar” allow you to save cookies set by the server in a request, and send them back on consecutive requests transparently. This allows, for example, going through an authentication request before sending the actual data-fetching request.

' See: http://framework.zend.com/manual/current/en/modules/zend.http.client.advanced.html#sending-multiple-requests-with-the-same-client


' Options

' KeepAlive

' If StoreCookies is enabled, the request object will automatically add cookies to the jar
' Used to manage and retain cookies between requests

' StoreCookies

' Methods:

' There needs to be come kind of way to set default headers and headers for specific-requests only
' DefaultHeader
' DefaultHeaders



Private Sub Class_Initialize()

End Sub

' Reuse HTTP Connections through a 'Connection' header.
Public Property Get KeepAlive()
	KeepAlive = pKeepAlive
End Property

Public Property Let KeepAlive(blnKeepAlive)
	pKeepAlive = blnKeepAlive
End Property

' If false, the received cookies as part of the HTTP response would be ignored.
Public Property Get StoreCookies()
	StoreCookies = pStoreCookies
End Property

Public Property Let StoreCookies(blnStoreCookies)
	pStoreCookies = blnStoreCookies
End Property

' Sends a DELETE request. Returns Response object.
Public Function DeleteReq()

End Function

' Sends a GET request. Returns Response object.
Public Function GetReq()

End Function

' Sends a HEAD request. Returns Response object.
Public Function HeadReq()

End Function

' Sends a OPTIONS request. Returns Response object.
Public Function OptionsReq()

End Function

' Sends a PATCH request. Returns Response object.
Public Function PatchReq()

End Function

' Sends a POST request. Returns Response object.
Public Function PostReq()

End Function

' Sends a PUT request. Returns Response object.
Public Function PutReq()

End Function

' Constructs and sends a Request. Returns Response object.
Public Function Request()

End Function

Public Sub Send()

End Sub

' Sessions can also be used to provide default data to the request methods. This
' is done by providing data to the properties on a Session object.
' Any dictionaries that you pass to a request method will be merged with the session-level
' values that are set. The method-level parameters override session parameters.
' Note, however, that method-level parameters will not be persisted across requests, even
' if using a session.
' Default Request Headers
' You can set default headers that will be sent on every request:

' Unirest\Request::defaultHeader("Header1", "Value1");
' Unirest\Request::defaultHeader("Header2", "Value2");
Public Sub DefaultHeader()

End Sub

' You can set default headers in bulk by passing an array:

' Unirest\Request::defaultHeaders(array(
	'	"Header1" => "Value1",
	'	"Header2" => "Value2"
' ));
Public Sub DefaultHeaders(arrHeaders)

End Sub

Public Sub ClearDefaultHeaders()

End Sub

Private Sub Class_Terminate()

End Sub
