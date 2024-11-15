Attribute VB_Name = "HaversineDist"
Public Function getDistances(latitude1 As Double, longitude1 As Double, latitude2 As Range, longitude2 As Range) As Double
    Dim earth_radius As Double
    Dim Pi As Double
    Dim deg2rad As Double
    Dim dLat As Double
    Dim dLon As Double
    Dim a As Double
    Dim c As Double
    Dim d As Double
    Dim minDistance As Double
    Dim i As Long

    earth_radius = 6371 ' Radius of the Earth in kilometers
    Pi = 3.14159265
    deg2rad = Pi / 180

    minDistance = 999999 ' Initialize minDistance to a large number

    ' Loop through the input ranges
    For i = 1 To latitude2.Count
        Dim currentLatitude As Double
        Dim currentLongitude As Double
        
        currentLatitude = latitude2(i)
        currentLongitude = longitude2(i)

        ' Ensure latitude and longitude are numeric and within valid ranges
        If IsNumeric(currentLatitude) And IsNumeric(currentLongitude) Then
            ' Convert to radians
            Dim lat1Rad As Double
            Dim lat2Rad As Double
            Dim lon1Rad As Double
            Dim lon2Rad As Double
            
            lat1Rad = latitude1 * deg2rad
            lat2Rad = currentLatitude * deg2rad
            lon1Rad = longitude1 * deg2rad
            lon2Rad = currentLongitude * deg2rad
            
            ' Calculate the differences
            dLat = lat2Rad - lat1Rad
            dLon = lon2Rad - lon1Rad

            ' Haversine formula calculation
            a = Sin(dLat / 2) * Sin(dLat / 2) + Cos(lat1Rad) * Cos(lat2Rad) * Sin(dLon / 2) * Sin(dLon / 2)
            c = 2 * WorksheetFunction.Asin(Sqr(a))

            d = earth_radius * c

            ' Update minDistance if the current distance is less
            If d < minDistance Then
                minDistance = d
            End If
        End If
    Next i

    ' Return the minimum distance found
    getDistances = minDistance
End Function

