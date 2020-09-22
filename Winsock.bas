Attribute VB_Name = "Winsock"
'Visual Basic Tutorial:
'The Winsock Control
'By: |P|h|r|o|z|e|n| Entertainment - Joker

'I love to program with Winsock, it is in my opinion, one of the most powerful extra controls Microsoft came up with. It is the lowest level network programming protocol, and lets you do anything. There is a catch to that of course, which is how hard it is to program it. This tutorial should get you up and running with using the Winsock Visual Basic Control. Also, I have included some extremely useful Functions for use with the Winsock Control.

'Ok to start us off with this tutorial, open a new Normal project and add the Winsock control to your project.


'--------------------------------------------------------------------------------

'Useful Winsock Functions:

'Winsock1.Listen 'This function makes your Winsock control listen on the LocalPort specified either with the function below or in the properties for this control.

'Winsock1.LocalPort 'This is the port where all incoming data arrives, including connection requests

'Winsock1.Connect [HostIP], [RemotePort] 'This will attempt to connect to the HostIP on the RemotePort of that computer, which should be what that computer is listening on. Both parameters are optional.

'Winsock1.Accept [requestID] 'This is how you accept a connection in the Winsock_ConnectionRequest Event. You must put Winsock1.Accept requestID.

'Winsock1.SendData [Data] 'This is what you use to send data with Winsock, you can send any time, but we recommend Strings with a special Function to Process it. The only 2 logical data types to send would be a Byte Array or a String.

'Winsock1.GetData [Data], [Type], [MaxLen] 'This is how you recieve all data from the Winsock control, Data = a variable, and Type = vbString,vbInteger etc. MaxLen is optional. This can only be used if data is sitting in the buffer, for instance using it in the Winsock_DataArrival event is a very practical use.

'Winsock1.RemoteHostIP [IPAddress] 'Use this to specify an IP address to connect to ahead of time.




'--------------------------------------------------------------------------------

'Winsock Terminology

'Port - A port is a number used to divide up all the connections going on in your computer. For instance, this browser is running on the HTML port, some number like 250, and any data send to or from goes to port 250 of the receiving computer.
'Listen - To listen means to set aside a certain port number, and wait for a connection request on that port.
'IP Address - This number is used on the internet kind of like a social security number for your computer, computers send data to your IP Address and the data is only recieved by you.

'--------------------------------------------------------------------------------

'Philosophies (How it Works)

'Alright, so you understand all the events and methods right? Now lets move on to the actual programming parts. This is where we get into Client/Server programming, which means you have a Client, for instance, this browser is connected to my server, and is a client. My server, on the other hand is connected to you, was waiting for you to connect and did connect you so you could read this, now it sent you this text when you requested it.

'Server Functions:
'--->Listen for users and connect them to the application
'--->Get data from users and return requested data
'--->Hold all application logic and process to return data

'Client Functions:
'--->Connect to server application
'--->Request data, recieve returned business logic and read

'Ok we 'll start with the server, since it is the backbone of the whole data sending process, you can either have one set server, that you will be operating and maintaining, or have an application with both Server and Client built into it, so anyone can host a server. Games and chat programs use the latter, but it is not often used for anything else. For bigger programs, you'll want to connect with more that 1 user, which requires more than 1 Winsock Control. So create 2 winsock controls, one called sckConnect, and one called sckConnection and set its index to 0.

'Now you can either add in my MakeListen() function, or create your own, but in the end, you need to make your server listen (sckConnect.Listen) for connections. Then, in the sckConnect_ConnectionRequest Event, make it accept the port with a sckConnection(num) Winsock and start data flinging on those connections. I highly suggest using our ProcessData() function for sending complex data because it makes it very easy to send 3 strings and decode them.

'Now on to the Client. The client is where your users will be, so it needs to look sleek, and have no performance flaws. Keep the sending and winsock processes away from them, so they can't mess it up. First, you'll want to connect to the server, on whatever RemotePort its on, by using sckConnect.Connect "TheServerIPAddres"

'Next, you want to send data, so make some Command Buttons send data to the Server, which in turn should suck it up and process it, and return something to know that it was processed. This is how Winsock should work, send a piece, and recieve a conformation or the requested data as a conformation.

'All data you recieve will be in the Winsock_DataArrival Event, so you should always use this. Usually, its common to make a string, use Winsock.GetData TheString, vbString to grab the data from winsock, and process it therefore grabbing all data when it is recieved.

'You can debug your Winsock applications on one computer, simply run two instances of Visual Basic and run both of them. If you have a modem, use the IP "127.0.0.1" or if you have a network card, use your network IP. You should test it on other computers before releasing to the public though, because sometimes its different.


'--------------------------------------------------------------------------------


  
