'
' SPDX-FileCopyrightText: 2020-2023 DB Systel GmbH
' SPDX-FileCopyrightText: 2023 Frank Schwab
'
' SPDX-License-Identifier: Apache-2.0
'
' Licensed under the Apache License, Version 2.0 (the "License");
' You may not use this file except in compliance with the License.
'
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'
' Author: Frank Schwab
'
' Version: 2.0.0
'
' Change history:
'    2020-04-21: V1.0.0: Created.
'    2023-08-13: V2.0.0: Much simpler redesign.
'

Option Strict On
Option Explicit On

''' <summary>
''' Converts integers from and to an unsigned packed byte array
''' </summary>
Public NotInheritable Class CompressedInteger
#Region "Private constants"
   '******************************************************************
   ' Private constants
   '******************************************************************

   '
   ' Constant for calculation
   '

   '
   ' It is not possible to write a byte constant in VB.
   '
   Private Const ONE_AS_BYTE As Byte = 1US

   Private Const OFFSET_VALUE As Byte = &H40

   Private Const MAX_ALLOWED_INTEGER As Integer = &H40404040 - 1

   ' Constants for masks

   Private Const NO_LENGTH_MASK_FOR_BYTE As Byte = OFFSET_VALUE - 1
   Private Const BYTE_MASK_FOR_INTEGER As Integer = &HFF

   ' Array manipulation constants

   Private Const RESULT_ARRAY_LENGTH As Integer = 4
   Private Const RESULT_MAX_INDEX As Integer = RESULT_ARRAY_LENGTH - 1
   Private Const LENGTH_BITS_SHIFT_VALUE As Integer = 6
#End Region

#Region "Public methods"
   '******************************************************************
   ' Public methods
   '******************************************************************

   ''' <summary>
   ''' Convert an integer into a packed unsigned integer byte array
   ''' </summary>
   ''' <remarks>
   ''' Valid integers range from 0 to 1,077,952,575.
   ''' All other numbers throw an <see cref="ArgumentException"/>.
   ''' </remarks>
   ''' <param name="anInteger">Number to convert to a packed unsigned integer</param>
   ''' <exception cref="ArgumentException">Thrown if <paramref name="anInteger"/> has not a value between 0 and 1,077,952,575 (inclusive)</exception>
   ''' <returns>packed unsigned integer byte array with integer as value</returns>
   Public Shared Function FromInteger(anInteger As Integer) As Byte()
      If anInteger < 0 Then _
         Throw New ArgumentException("Integer must not be negative")

      If anInteger > MAX_ALLOWED_INTEGER Then _
         Throw New ArgumentException("Integer too large for packed integer")

      Dim result As Byte() = New Byte(0 To RESULT_MAX_INDEX) {}

      Dim temp As Integer = anInteger
      Dim actIndex As Integer = RESULT_MAX_INDEX

      While temp >= OFFSET_VALUE
         Dim b As Integer = temp And BYTE_MASK_FOR_INTEGER

         temp >>= 8

         If b >= OFFSET_VALUE Then
            b -= OFFSET_VALUE
         Else
            b += 256 - OFFSET_VALUE
            temp -= 1
         End If

         result(actIndex) = CByte(b)

         actIndex -= 1
      End While

      'Add length flag
      result(actIndex) = CByte(temp Or (RESULT_MAX_INDEX - actIndex) << LENGTH_BITS_SHIFT_VALUE)

      Return ArrayHelper.CopyOf(result, actIndex, RESULT_ARRAY_LENGTH - actIndex)
   End Function

   ''' <summary>
   ''' Convert a packed unsigned integer byte array in a possibly larger array to an integer.
   ''' </summary>
   ''' <param name="arrayWithPackedNumber">Array in which the packed unsigned integer byte array resides.</param>
   ''' <param name="startIndex">Start index of packed unsigned integer byte array in the byte array.</param>
   ''' <returns>Converted integer (value between 0 and 1,077,952,575).</returns>
   Public Shared Function ToInteger(arrayWithPackedNumber As Byte(), startIndex As Integer) As Integer
      RequireNonNull(arrayWithPackedNumber, NameOf(arrayWithPackedNumber))

      Dim arrayLength As Integer = arrayWithPackedNumber.Length

      If arrayLength = 0 Then _
         Throw New ArgumentException("Array must have a length greater 0")

      Dim expectedLength As Integer = GetExpectedLengthWithoutCheck(arrayWithPackedNumber, startIndex)

      If startIndex + expectedLength > arrayLength Then _
         Throw New ArgumentException("Array is too short for packed number")

      ' Decompress the byte array
      Dim temp As Integer = arrayWithPackedNumber(startIndex) And NO_LENGTH_MASK_FOR_BYTE

      For i As Integer = startIndex + 1 To startIndex + expectedLength - 1
         temp = ((temp << 8) Or arrayWithPackedNumber(i)) + OFFSET_VALUE
      Next i

      Return temp
   End Function

   ''' <summary>
   ''' Convert a packed unsigned integer byte array into an integer.
   ''' </summary>
   ''' <param name="aPackedUnsignedInteger">Packed unsigned integer byte array.</param>
   ''' <exception cref="ArgumentException">Thrown if the actual length of the packed number does not match the expected length.</exception>
   ''' <returns>Converted integer (value between 0 and 1,077,952,575).</returns>
   Public Shared Function ToInteger(aPackedUnsignedInteger As Byte()) As Integer
      Return ToInteger(aPackedUnsignedInteger, 0)
   End Function

   ''' <summary>
   ''' Get expected length of packed unsigned integer byte array in a a possibly larger array.
   ''' </summary>
   ''' <param name="arrayWithPackedNumber">Array in which the packed unsigned integer byte array resides.</param>
   ''' <param name="startIndex">Start index of packed unsigned integer byte array in the byte array.</param>
   ''' <returns>Expected length (1 to 4)</returns>
   Public Shared Function GetExpectedLength(arrayWithPackedNumber As Byte(), startindex As Integer) As Byte
      RequireNonNull(arrayWithPackedNumber, NameOf(arrayWithPackedNumber))

      Return GetExpectedLengthWithoutCheck(arrayWithPackedNumber, startindex)
   End Function

   ''' <summary>
   ''' Get expected length of packed unsigned integer byte array from first byte.
   ''' </summary>
   ''' <param name="aPackedUnsignedInteger">Packed unsigned integer byte array.</param>
   ''' <returns>Expected length (1 to 4)</returns>
   Public Shared Function GetExpectedLength(aPackedUnsignedInteger As Byte()) As Byte
      Return GetExpectedLength(aPackedUnsignedInteger, 0)
   End Function

   ''' <summary>
   ''' Convert a decimal byte array that is supposed to be a packed unsigned integer
   ''' into a string.
   ''' </summary>
   ''' <param name="aPackedUnsignedInteger">Byte array of packed unsigned integer</param>
   ''' <exception cref="ArgumentNullException">Thrown if <paramref name="aPackedUnsignedInteger"/> is <c>Nothing</c>.</exception>
   ''' <returns>String representation of the given packed unsigned integer</returns>
#Disable Warning BC40005 ' Member shadows an overridable method in the base type: This does *not* override Object.ToString()
   Public Shared Function ToString(aPackedUnsignedInteger As Byte()) As String
#Enable Warning BC40005 ' Member shadows an overridable method in the base type
      Return ToInteger(aPackedUnsignedInteger).ToString()
   End Function
#End Region

#Region "Private methods"

#Region "Internal calculation methods"
   ''' <summary>
   ''' Get expected length of packed unsigned integer byte array in a a possibly larger array.
   ''' </summary>
   ''' <remarks>This method does not check if the supplied array is <c>Nothing</c> as it assumes that this check has already been made.</remarks>
   ''' <param name="arrayWithPackedNumber">Array in which the packed unsigned integer byte array resides.</param>
   ''' <param name="startIndex">Start index of packed unsigned integer byte array in the byte array.</param>
   ''' <returns>Expected length (1 to 4)</returns>
   Private Shared Function GetExpectedLengthWithoutCheck(arrayWithPackedNumber As Byte(), startindex As Integer) As Byte
      Return (arrayWithPackedNumber(startindex) >> 6) + ONE_AS_BYTE
   End Function
#End Region

#Region "Exception helper methods"
   ''' <summary>
   ''' Check if object is null and throw an exception, if it is.
   ''' </summary>
   ''' <param name="anObject">Object to check.</param>
   ''' <param name="parameterName">Parameter name for exception.</param>
   ''' <exception cref="ArgumentNullException">Thrown when <paramref name="anObject"/> is <c>Nothing</c>.</exception>
   Private Shared Sub RequireNonNull(anObject As Object, parameterName As String)
      If anObject Is Nothing Then _
         Throw New ArgumentNullException(parameterName)
   End Sub
#End Region
#End Region
End Class
