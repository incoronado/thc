Option Explicit
'On Error Resume Next

Dim Action
Dim SleepVar

SleepVar = 5

Do
	Sleep SleepVar
		
	Action = GetPropertyValue ("Weather.Action")
	If Action <> "Idle" Then
		SetpropertyValue "Weather.Action", "Idle"
		Sleep SleepVar
		If Action <> "" Then
			Call ProcessMessage(Action)
		End If
	End If
Loop

Sub ProcessMessage(Action)
    'MsgBox "Ready Command"
    Select Case Action 
	' Configuration Command
		Case "GetCurrentWeather"
	     GetCurrentWeather GetPropertyValue("Weather.Weather-zipcode")
		Case "GetWeatherForecast"
		 GetWeatherForecast GetPropertyValue("Weather.Weather-zipcode") 		
	   End Select
End Sub


' Current Weather
' lattitude
' longitude
' elevation
' feelslike_f
' visibility_mi
' dewpoint_f
' cloudy
' weather
' wind_dir
' wind_degrees
' wind_mph
' wind_gust_mph
' pressure_in
' pressure_trend
' windchill_f
' precip_today_in
' precip_1hr_in
' windchill_f
' UV
' relative_humidity

' Weather Forecast
'period
'icon
'icon_url
'title
'fctext
'fcttext_metric
'pop
'date.epoch
'date.pretty
'date.day
'date.month
'date.year
'date.yday
'date.hour
'date.min
'date.sec
'date.isdst
'date.monthname
'date.weekday_short
'date.weekday
'date.ampm
'date.tz_short
'date.tz_long
'period
'high
'low
'conditions
'icon
'icon_url
'skyicon
'pop
'qpf_allday
'qpf_day
'qpf_night
'snow_allday
'snow_day
'snow_night
'maxwind
'avewind
'avehumidity
'maxhumidity
'minhumidity



'    [Light/Heavy] Drizzle
'    [Light/Heavy] Rain
'    [Light/Heavy] Snow
'    [Light/Heavy] Snow Grains
'    [Light/Heavy] Ice Crystals
'    [Light/Heavy] Ice Pellets
'    [Light/Heavy] Hail
'    [Light/Heavy] Mist
'    [Light/Heavy] Fog
'    [Light/Heavy] Fog Patches
'    [Light/Heavy] Smoke
'    [Light/Heavy] Volcanic Ash
'    [Light/Heavy] Widespread Dust
'    [Light/Heavy] Sand
'    [Light/Heavy] Haze
'    [Light/Heavy] Spray
'    [Light/Heavy] Dust Whirls
'    [Light/Heavy] Sandstorm
'    [Light/Heavy] Low Drifting Snow
'    [Light/Heavy] Low Drifting Widespread Dust
'    [Light/Heavy] Low Drifting Sand
'    [Light/Heavy] Blowing Snow
'    [Light/Heavy] Blowing Widespread Dust
'    [Light/Heavy] Blowing Sand
'    [Light/Heavy] Rain Mist
'    [Light/Heavy] Rain Showers
'    [Light/Heavy] Snow Showers
'    [Light/Heavy] Snow Blowing Snow Mist
'    [Light/Heavy] Ice Pellet Showers
'    [Light/Heavy] Hail Showers
'    [Light/Heavy] Small Hail Showers
'    [Light/Heavy] Thunderstorm
'    [Light/Heavy] Thunderstorms and Rain
'    [Light/Heavy] Thunderstorms and Snow
'    [Light/Heavy] Thunderstorms and Ice Pellets
'    [Light/Heavy] Thunderstorms with Hail
'    [Light/Heavy] Thunderstorms with Small Hail
'    [Light/Heavy] Freezing Drizzle
'    [Light/Heavy] Freezing Rain
'    [Light/Heavy] Freezing Fog
'    Patches of Fog
'    Shallow Fog
'   Partial Fog
'    Overcast
'    Clear
'    Partly Cloudy
'    Mostly Cloudy
'    Scattered Clouds
'    Small Hail
'    Squalls
'    Funnel Cloud
'    Unknown Precipitation
'    Unknown


Sub GetWeatherForecast(zipcode)
    Dim arr(8,1), url, xmlDoc, ColItem, i, objItem, cnt
	'url = "http://api.wunderground.com/api/7edbcd9d2f86e37b/forecast/q/ZIP/" & zipcode & ".xml"
	url = "http://api.wunderground.com/api/7edbcd9d2f86e37b/forecast/q/CA/San Diego.xml"

	Set xmlDoc = createobject("Microsoft.XMLDOM")
	xmlDoc.async = "false"
	xmlDoc.load (url)


	arr(0,0) = "/response/forecast/simpleforecast/forecastdays/forecastday/high/fahrenheit"
	arr(0,1) = "high"
	arr(1,0) = "/response/forecast/simpleforecast/forecastdays/forecastday/low/fahrenheit"
	arr(1,1) = "low"
	arr(2,0) = "/response/forecast/simpleforecast/forecastdays/forecastday/pop"
	arr(2,1) = "pop"
	arr(3,0) = "/response/forecast/simpleforecast/forecastdays/forecastday/qpf_allday/in"
	arr(3,1) = "qpf_allday"
	arr(4,0) = "/response/forecast/simpleforecast/forecastdays/forecastday/conditions"
	arr(4,1) = "conditions"
	arr(5,0) = "/response/forecast/simpleforecast/forecastdays/forecastday/avehumidity"
	arr(5,1) = "avehumidity"
	arr(6,0) = "/response/forecast/simpleforecast/forecastdays/forecastday/date/weekday"
	arr(6,1) = "weekday"
	arr(7,0) = "/response/forecast/simpleforecast/forecastdays/forecastday/avewind/dir"
	arr(7,1) = "wind_dir"
	arr(8,0) = "/response/forecast/simpleforecast/forecastdays/forecastday/avewind/mph"
	arr(8,1) = "wind_mph"

	For i = 0 To 8
		Set colItem = xmlDoc.selectNodes(arr(i,0))
		cnt=1
		For Each objItem in colItem
			 if arr(i,1) = "high" Or arr(i,1) = "low" Then
				SetpropertyValue "Weather.Forecast" & cnt & "-" & arr(i,1), CInt(objItem.text)  & chr(176)
			Else
				SetpropertyValue "Weather.Forecast" & cnt & "-" & arr(i,1), objItem.text
			End if	
			cnt=cnt+1
		Next
	
		'WScript.Echo "XML Element Count :" & xmlCol.length
	Next
	
	Set colItem = Nothing	
	Set xmlDoc = Nothing


End Sub

Sub GetCurrentWeather(zipcode)
   Dim arr(13,1), url, xmlDoc, ColItem, i, objItem
	url = "http://api.wunderground.com/api/7edbcd9d2f86e37b/conditions/q/ZIP/" & zipcode & "92118.xml"

	Set xmlDoc = createobject("Microsoft.XMLDOM")
	xmlDoc.async = "false"
	xmlDoc.load (url)


	arr(0,0) = "/response/current_observation/display_location/full"
	arr(0,1) = "location"
	arr(1,0) = "/response/current_observation/display_location/longitude"
	arr(1,1) = "longitude"
	arr(2,0) = "/response/current_observation/display_location/latitude"
	arr(2,1) = "latitude"
	arr(3,0) = "/response/current_observation/observation_time"
	arr(3,1) = "observation_time"
	arr(4,0) = "/response/current_observation/weather"
	arr(4,1) = "weather"
	arr(5,0) = "/response/current_observation/temp_f"
	arr(5,1) = "temp"
	arr(6,0) = "/response/current_observation/relative_humidity"
	arr(6,1) = "humidity"
	arr(7,0) = "/response/current_observation/wind_mph"
	arr(7,1) = "wind_mph"
	arr(8,0) = "/response/current_observation/wind_dir"
	arr(8,1) = "wind_dir"
	arr(9,0) = "/response/current_observation/dewpoint_f"
	arr(9,1) = "dewpoint_f"
	arr(10,0) = "/response/current_observation/precip_today_in"
	arr(10,1) = "precip_today_in"
	arr(11,0) = "/response/current_observation/pressure_trend"
	arr(11,1) = "pressure_trend"
	arr(12,0) = "/response/current_observation/visibility_mi"
	arr(12,1) = "visibility_mi"
	arr(13,0) = "/response/current_observation/pressure_in"
	arr(13,1) = "pressure_in"
	
 	
	For i = 0 To 13
		Set colItem = xmlDoc.selectNodes(arr(i,0))
		For Each objItem in colItem
		    if arr(i,1) = "temp" Then
				SetpropertyValue "Weather.Conditions-" & arr(i,1), CInt(objItem.text) & chr(176)
			Else
				SetpropertyValue "Weather.Conditions-" & arr(i,1), objItem.text
			End if	
		Next
		'WScript.Echo "XML Element Count :" & xmlCol.length
	Next
	
	Set colItem = Nothing	
	Set xmlDoc = Nothing

End Sub

