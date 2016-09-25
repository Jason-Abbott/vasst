<%
	Dim states
	states = Array("-- UNITED STATES --", "[AL] Alabama", "[AK] Alaska", "[AZ] Arizona", "[AR] Arkansas", "[CA] California", "[CO] Colorado", "[CT] Connecticut", "[DC] District of Columbia", "[DE] Delaware", "[FL] Florida", "[GA] Georgia", "[HI] Hawaii", "[ID] Idaho", "[IL] Illinois", "[IN] Indiana", "[IA] Iowa", "[KS] Kansas", "[KY] Kentucky", "[LA] Louisiana", "[ME] Maine", "[MD] Maryland", "[MA] Massachusetts", "[MI] Michigan", "[MN] Minnesota", "[MS] Mississippi", "[MO] Missouri", "[MT] Montana", "[NE] Nebraska", "[NV] Nevada", "[NH] New Hampshire", "[NJ] New Jersey", "[NM] New Mexico", "[NY] New York", "[NC] North Carolina", "[ND] North Dakota", "[OH] Ohio", "[OK] Oklahoma", "[OR] Oregon", "[PA] Pennsylvania", "[RI] Rhode Island", "[SC] South Carolina", "[SD] South Dakota", "[TN] Tennessee", "[TX] Texas", "[UT] Utah", "[VT] Vermont", "[VA] Virginia", "[WA] Washington", "[WV] West Virginia", "[WI] Wisconsin", "[WY] Wyoming", "-- CANADA --", "[AB] Alberta", "[BC] British Columbia", "[MB] Manitoba", "[NB] New Brunswick", "[NF] Newfoundland and Labrador", "[NT] Northwest Territories", "[NS] Nova Scotia", "[NU] Nunavut", "[ON] Ontario", "[PE] Prince Edward Island", "[PQ] Quebec", "[SK] Saskatchewan", "[YT] Yukon Territory")
%>
			<select size="1" name="<%=strFormName%>">
				<option value="">Select a state...</option>
				<% 
					For Each state In states
						If (InStr(state, "[") < InStr(state, "]")) Then
							stateAbbr = Right(Left(state, 3), 2)
						Else
							stateAbbr = ""
						End If
						
						If ((Trim(Lcase(stateAbbr)) = Trim(Lcase(Request.Form(strFormName)))) And (stateAbbr <> "")) Then
							thisSelected = " selected"
						Else
							thisSelected = ""
						End If
						print("<option value=""" & stateAbbr & """" & thisSelected & ">" & state & "</option>")
					Next
				%>
			</select>