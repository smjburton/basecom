Option Explicit

Class base_IO_MemoryStream
	Private p_MemoryStream

	Private Sub Class_Initialize()
		Set p_MemoryStream = CreateObject("System.IO.MemoryStream")
	End Sub


	' Properties


	Public Property Get CanRead()
		CanRead = p_MemoryStream.CanRead 
	End Property

	Public Property Get CanSeek()
		CanSeek = p_MemoryStream.CanSeek 
	End Property

	Public Property Get CanTimeout()
		CanTimeout = p_MemoryStream.CanTimeout 
	End Property

	Public Property Get CanWrite()
		CanWrite = p_MemoryStream.CanWrite 
	End Property

	Public Property Get Capacity()
		Capacity = p_MemoryStream.Capacity 
	End Property

	Public Property Let Capacity(intCapacity)
		p_MemoryStream.Capacity = intCapacity
	End Property

	Public Property Get Length()
		Length = p_MemoryStream.Length 
	End Property

	Public Property Get Position()
		Position = p_MemoryStream.Position 
	End Property

	Public Property Let Position(lngPosition)
		p_MemoryStream.Position = lngPosition
	End Property

	Public Property Get ReadTimeout()
		ReadTimeout = p_MemoryStream.ReadTimeout 
	End Property

	Public Property Let ReadTimeout(intReadTimeout)
		p_MemoryStream.ReadTimeout = intReadTimeout
	End Property

	Public Property Get WriteTimeout()
		WriteTimeout = p_MemoryStream.WriteTimeout 
	End Property

	Public Property Let WriteTimeout(intWriteTimeout)
		p_MemoryStream.WriteTimeout = intWriteTimeout 
	End Property


	' Methods


	Public Function BeginRead(arrByteBuffer(), intOffset, intCount, objAsyncCallback, objState)
		Set BeginRead = p_MemoryStream.BeginRead(arrByte(), intOffset, intCount, objAsyncCallback, objState)
	End Function

	Public Function BeginWrite(arrByteBuffer(), intOffset, intCount, objAsyncCallback, objState)
		Set BeginWrite = p_MemoryStream.BeginWrite(arrByteBuffer(), intOffset, intCount, objAsyncCallback, objState)
	End Function

	Public Sub Close()
		p_MemoryStream.Close
	End Sub

	Public Sub CopyTo(objStream) ' Optional params: CopyTo(objStream, intBufferSize)
		p_MemoryStream.CopyTo objStream
	End Sub

	Public Function CopyToAsync(objStream) ' Optional params: CopyToAsync(objStream, intBufferSize); CopyToAsync(objStream, intBufferSize, objCancellationToken)
		Set CopyToAsync = p_MemoryStream.CopyToAsync(objStream)
	End Function

	Public Function CreateObjRef(objType)
		Set CreateObjRef = p_MemoryStream.CreateObjRef(objType)
	End Function

	Public Sub Dispose() ' Optional params: Dispose(blnDisposing)
		p_MemoryStream.Dispose
	End Sub

	Public Function EndRead(objAsyncResult)
		EndRead = p_MemoryStream.EndRead(objAsyncResult)
	End Function

	Public Sub EndWrite(objAsyncResult)
		p_MemoryStream.EndWrite objAsyncResult
	End Sub

	Public Function Equals(objObject)
		Equals = p_MemoryStream.Equals(objObject)
	End Function

	Public Sub Finalize()
		p_MemoryStream.Finalize
	End Sub

	Public Sub Flush()
		p_MemoryStream.Flush
	End Sub

	Public Function FlushAsync() ' Optional params: FlushAsync(objCancellationToken)
		Set FlushAsync = p_MemoryStream.FlushAsync()
	End Function

	Public Function GetBuffer()
		GetBuffer = p_MemoryStream.GetBuffer()
	End Function

	Public Function GetHashCode()
		GetHashCode = p_MemoryStream.GetHashCode()
	End Function

	Public Function GetLifetimeService()
		Set GetLifetimeService = p_MemoryStream.GetLifetimeService()
	End Function

	Public Function GetType()
		Set GetType = p_MemoryStream.GetType()
	End Function

	Public Function InitializeLifetimeService()
		Set InitializeLifetimeService = p_MemoryStream.InitializeLifetimeService()
	End Function

	Public Function MemberwiseClone() ' Optional params: MemberwiseClone(Boolean)
		Set MemberwiseClone = p_MemoryStream.MemberwiseClone()
	End Function

	Public Function Read(arrByteBuffer(), intOffset, intCount)
		Read = p_MemoryStream.Read(arrByteBuffer(), intOffset, intCount)
	End Function

	Public Function ReadAsync(arrByteBuffer(), intOffset, intCount) ' Optional params: ReadAsync(arrByteBuffer(), intOffset, intCount, objCancellationToken)
		Set ReadAsync = p_MemoryStream.ReadAsync(arrByteBuffer(), intOffset, intCount)
	End Function

	Public Function ReadByte()
		ReadByte = p_MemoryStream.ReadByte()
	End Function

	Public Function Seek(lngOffset, objSeekOrigin)
		Seek = p_MemoryStream.Seek(lngOffset, objSeekOrigin)
	End Function

	Public Sub SetLength(lngValue)
		p_MemoryStream.SetLength lngValue
	End Sub

	Public Function ToArray()
		ToArray = p_MemoryStream.ToArray()
	End Function

	Public Function ToString()
		ToString = p_MemoryStream.ToString()
	End Function

	Public Function TryGetBuffer(objArraySegment)
		TryGetBuffer = p_MemoryStream.TryGetBuffer(objArraySegment)
	End Function

	Public Sub Write(arrByteBuffer(), intOffset, intCount)
		p_MemoryStream.Write arrByteBuffer(), intOffset, intCount 
	End Sub

	Public Function WriteAsync(arrByteBuffer(), intOffset, intCount) ' Optional params: WriteAsync(Byte(), Int32, Int32, CancellationToken)
		Set WriteAsync = p_MemoryStream.WriteAsync(arrByteBuffer(), intOffset, intCount)
	End Function

	Public Sub WriteByte(objByte)
		p_MemoryStream.WriteByte objByte
	End Sub

	Public Sub WriteTo(objStream)
		p_MemoryStream.WriteTo objStream
	End Sub

	Private Sub Class_Terminate()
		Set p_MemoryStream = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_IO_MemoryStream.vbs" Then

End If
