<%
'#################################################################################
'## Snitz Forums 2000 v3.4.07
'#################################################################################
'## Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen,
'##                       Huw Reddick and Richard Kinser
'##
'## This program is free software; you can redistribute it and/or
'## modify it under the terms of the GNU General Public License
'## as published by the Free Software Foundation; either version 2
'## of the License, or (at your option) any later version.
'##
'## All copyright notices regarding Snitz Forums 2000
'## must remain intact in the scripts and in the outputted HTML
'## The "powered by" text/logo with a link back to
'## http://forum.snitz.com in the footer of the pages MUST
'## remain visible when the pages are viewed on the internet or intranet.
'##
'## This program is distributed in the hope that it will be useful,
'## but WITHOUT ANY WARRANTY; without even the implied warranty of
'## MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'## GNU General Public License for more details.
'##
'## You should have received a copy of the GNU General Public License
'## along with this program; if not, write to the Free Software
'## Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'##
'## Support can be obtained from our support forums at:
'## http://forum.snitz.com
'##
'## Correspondence and Marketing Questions can be sent to:
'## manderson@snitz.com
'##
'#################################################################################

	Response.Write "<option value=""Afghanistan"">Afghanistan</option>" & strLE & _
			"<option value=""Albania"">Albania</option>" & strLE & _
			"<option value=""Algeria"">Algeria</option>" & strLE & _
			"<option value=""Andorra"">Andorra</option>" & strLE & _
			"<option value=""Angola"">Angola</option>" & strLE & _
			"<option value=""Anguilla"">Anguilla</option>" & strLE & _
			"<option value=""Antigua and Barbuda"">Antigua and Barbuda</option>" & strLE & _
			"<option value=""Argentina"">Argentina</option>" & strLE & _
			"<option value=""Armenia"">Armenia</option>" & strLE & _
			"<option value=""Aruba"">Aruba</option>" & strLE & _
			"<option value=""Australia"">Australia</option>" & strLE & _
			"<option value=""Austria"">Austria</option>" & strLE & _
			"<option value=""Azerbaijan"">Azerbaijan</option>" & strLE & _
			"<option value=""Azores"">Azores</option>" & strLE & _
			"<option value=""Bahamas"">Bahamas</option>" & strLE & _
			"<option value=""Bahrain"">Bahrain</option>" & strLE & _
			"<option value=""Bangladesh"">Bangladesh</option>" & strLE & _
			"<option value=""Barbados"">Barbados</option>" & strLE & _
			"<option value=""Belarus"">Belarus</option>" & strLE & _
			"<option value=""Belgium"">Belgium</option>" & strLE & _
			"<option value=""Belize"">Belize</option>" & strLE & _
			"<option value=""Benin"">Benin</option>" & strLE & _
			"<option value=""Bermuda"">Bermuda</option>" & strLE & _
			"<option value=""Bhutan"">Bhutan</option>" & strLE & _
			"<option value=""Bolivia"">Bolivia</option>" & strLE & _
			"<option value=""Borneo"">Borneo</option>" & strLE & _
			"<option value=""Bosnia and Herzegovina"">Bosnia and Herzegovina</option>" & strLE & _
			"<option value=""Botswana"">Botswana</option>" & strLE & _
			"<option value=""Brazil"">Brazil</option>" & strLE & _
			"<option value=""British Indian Ocean Territories"">British Indian Ocean Territories</option>" & strLE & _
			"<option value=""Brunei"">Brunei</option>" & strLE & _
			"<option value=""Bulgaria"">Bulgaria</option>" & strLE & _
			"<option value=""Burkina Faso (Upper Volta)"">Burkina Faso (Upper Volta)</option>" & strLE & _
			"<option value=""Burundi"">Burundi</option>" & strLE & _
			"<option value=""Camaroon"">Camaroon</option>" & strLE & _
			"<option value=""Cambodia"">Cambodia</option>" & strLE & _
			"<option value=""Canada"">Canada</option>" & strLE & _
			"<option value=""Canary Islands"">Canary Islands</option>" & strLE & _
			"<option value=""Cape Vere Islands"">Cape Vere Islands</option>" & strLE & _
			"<option value=""Cayman Islands"">Cayman Islands</option>" & strLE & _
			"<option value=""Central African Rep"">Central African Rep</option>" & strLE & _
			"<option value=""Chad"">Chad</option>" & strLE & _
			"<option value=""Chile"">Chile</option>" & strLE & _
			"<option value=""China"">China</option>" & strLE & _
			"<option value=""Christmas Island"">Christmas Island</option>" & strLE & _
			"<option value=""Colombia"">Colombia</option>" & strLE & _
			"<option value=""Comoros Islands"">Comoros Islands</option>" & strLE & _
			"<option value=""Congo, Democratic Republic of"">Congo, Democratic Republic of</option>" & strLE & _
			"<option value=""Costa Rica"">Costa Rica</option>" & strLE & _
			"<option value=""Croatia"">Croatia</option>" & strLE & _
			"<option value=""Cuba"">Cuba</option>" & strLE & _
			"<option value=""Cyprus"">Cyprus</option>" & strLE & _
			"<option value=""Czech Republic"">Czech Republic</option>" & strLE & _
			"<option value=""Denmark"">Denmark</option>" & strLE & _
			"<option value=""Djibouti"">Djibouti</option>" & strLE & _
			"<option value=""Dominica"">Dominica</option>" & strLE & _
			"<option value=""Dominican Republic"">Dominican Republic</option>" & strLE & _
			"<option value=""East Timor"">East Timor</option>" & strLE & _
			"<option value=""Ecuador"">Ecuador</option>" & strLE & _
			"<option value=""Egypt"">Egypt</option>" & strLE & _
			"<option value=""El Salvador"">El Salvador</option>" & strLE & _
			"<option value=""Equatorial Guinea"">Equatorial Guinea</option>" & strLE & _
			"<option value=""Eritria"">Eritria</option>" & strLE & _
			"<option value=""Estonia"">Estonia</option>" & strLE & _
			"<option value=""Ethiopia"">Ethiopia</option>" & strLE & _
			"<option value=""Falkland Islands"">Falkland Islands</option>" & strLE & _
			"<option value=""Faroe Islands"">Faroe Islands</option>" & strLE & _
			"<option value=""Fed Rep Yugoslavia"">Fed Rep Yugoslavia</option>" & strLE & _
			"<option value=""Fiji"">Fiji</option>" & strLE & _
			"<option value=""Finland"">Finland</option>" & strLE & _
			"<option value=""France"">France</option>" & strLE & _
			"<option value=""French Guiana"">French Guiana</option>" & strLE & _
			"<option value=""French Polynesia"">French Polynesia</option>" & strLE & _
			"<option value=""Fyro Macedonia"">Fyro Macedonia</option>" & strLE & _
			"<option value=""Gabon"">Gabon</option>" & strLE & _
			"<option value=""Gambia"">Gambia</option>" & strLE & _
			"<option value=""Georgia"">Georgia</option>" & strLE & _
			"<option value=""Germany"">Germany</option>" & strLE & _
			"<option value=""Ghana"">Ghana</option>" & strLE & _
			"<option value=""Gibraltar"">Gibraltar</option>" & strLE & _
			"<option value=""Greece"">Greece</option>" & strLE & _
			"<option value=""Greenland"">Greenland</option>" & strLE & _
			"<option value=""Grenada"">Grenada</option>" & strLE & _
			"<option value=""Guadeloupe"">Guadeloupe</option>" & strLE & _
			"<option value=""Guatemala"">Guatemala</option>" & strLE & _
			"<option value=""Guinea"">Guinea</option>" & strLE & _
			"<option value=""Guinea-Bissau"">Guinea-Bissau</option>" & strLE & _
			"<option value=""Guyana"">Guyana</option>" & strLE & _
			"<option value=""Haiti"">Haiti</option>" & strLE & _
			"<option value=""Honduras"">Honduras</option>" & strLE & _
			"<option value=""Hong Kong"">Hong Kong</option>" & strLE & _
			"<option value=""Hungary"">Hungary</option>" & strLE & _
			"<option value=""Iceland"">Iceland</option>" & strLE & _
			"<option value=""India"">India</option>" & strLE & _
			"<option value=""Indonesia"">Indonesia</option>" & strLE & _
			"<option value=""Iran"">Iran</option>" & strLE & _
			"<option value=""Iraq"">Iraq</option>" & strLE & _
			"<option value=""Ireland"">Ireland</option>" & strLE & _
			"<option value=""Israel"">Israel</option>" & strLE & _
			"<option value=""Italy"">Italy</option>" & strLE & _
			"<option value=""Ivory Coast"">Ivory Coast</option>" & strLE & _
			"<option value=""Jamaica"">Jamaica</option>" & strLE & _
			"<option value=""Japan"">Japan</option>" & strLE & _
			"<option value=""Jordan"">Jordan</option>" & strLE & _
			"<option value=""Kazakhstan"">Kazakhstan</option>" & strLE & _
			"<option value=""Kenya"">Kenya</option>" & strLE & _
			"<option value=""Kiribati"">Kiribati</option>" & strLE & _
			"<option value=""Korea"">Korea</option>" & strLE & _
			"<option value=""Kuwait"">Kuwait</option>" & strLE & _
			"<option value=""Kyrgyzstan"">Kyrgyzstan</option>" & strLE & _
			"<option value=""Laos"">Laos</option>" & strLE & _
			"<option value=""Latvia"">Latvia</option>" & strLE & _
			"<option value=""Lebanon"">Lebanon</option>" & strLE & _
			"<option value=""Lesotho"">Lesotho</option>" & strLE & _
			"<option value=""Liberia"">Liberia</option>" & strLE & _
			"<option value=""Libya"">Libya</option>" & strLE & _
			"<option value=""Liechtenstein"">Liechtenstein</option>" & strLE & _
			"<option value=""Lithuania"">Lithuania</option>" & strLE & _
			"<option value=""Luxembourg"">Luxembourg</option>" & strLE & _
			"<option value=""Macao"">Macao</option>" & strLE & _
			"<option value=""Madagascar"">Madagascar</option>" & strLE & _
			"<option value=""Malawi"">Malawi</option>" & strLE & _
			"<option value=""Malaysia"">Malaysia</option>" & strLE & _
			"<option value=""Maldives"">Maldives</option>" & strLE & _
			"<option value=""Mali"">Mali</option>" & strLE & _
			"<option value=""Malta"">Malta</option>" & strLE & _
			"<option value=""Martinique"">Martinique</option>" & strLE & _
			"<option value=""Mauritania"">Mauritania</option>" & strLE & _
			"<option value=""Mauritius"">Mauritius</option>" & strLE & _
			"<option value=""Mexico"">Mexico</option>" & strLE & _
			"<option value=""Moldova"">Moldova</option>" & strLE & _
			"<option value=""Monaco"">Monaco</option>" & strLE & _
			"<option value=""Mongolia"">Mongolia</option>" & strLE & _
			"<option value=""Montserrat"">Montserrat</option>" & strLE & _
			"<option value=""Morocco"">Morocco</option>" & strLE & _
			"<option value=""Mozambique"">Mozambique</option>" & strLE & _
			"<option value=""Myanmar (Burma)"">Myanmar (Burma)</option>" & strLE & _
			"<option value=""Namibia"">Namibia</option>" & strLE & _
			"<option value=""Naura"">Naura</option>" & strLE & _
			"<option value=""Nepal"">Nepal</option>" & strLE & _
			"<option value=""Netherlands"">Netherlands</option>" & strLE & _
			"<option value=""Netherlands Antilles"">Netherlands Antilles</option>" & strLE & _
			"<option value=""New Caledonia"">New Caledonia</option>" & strLE & _
			"<option value=""New Zealand"">New Zealand</option>" & strLE & _
			"<option value=""Nicaragua"">Nicaragua</option>" & strLE & _
			"<option value=""Niger"">Niger</option>" & strLE & _
			"<option value=""Nigeria"">Nigeria</option>" & strLE & _
			"<option value=""Niue"">Niue</option>" & strLE & _
			"<option value=""Norway"">Norway</option>" & strLE & _
			"<option value=""Oman"">Oman</option>" & strLE & _
			"<option value=""Pakistan"">Pakistan</option>" & strLE & _
			"<option value=""Panama"">Panama</option>" & strLE & _
			"<option value=""Papua New Guinea"">Papua New Guinea</option>" & strLE & _
			"<option value=""Paraguay"">Paraguay</option>" & strLE & _
			"<option value=""Peru"">Peru</option>" & strLE & _
			"<option value=""Philippines"">Philippines</option>" & strLE & _
			"<option value=""Pitcairn Island"">Pitcairn Island</option>" & strLE & _
			"<option value=""Poland"">Poland</option>" & strLE & _
			"<option value=""Portugal"">Portugal</option>" & strLE & _
			"<option value=""Qatar"">Qatar</option>" & strLE & _
			"<option value=""Republic of Korea"">Republic of Korea</option>" & strLE & _
			"<option value=""Reunion Island"">Reunion Island</option>" & strLE & _
			"<option value=""Romania"">Romania</option>" & strLE & _
			"<option value=""Russia"">Russia</option>" & strLE & _
			"<option value=""Rwanda"">Rwanda</option>" & strLE & _
			"<option value=""Saint Barthelemy"">Saint Barthelemy</option>" & strLE & _
			"<option value=""Saint Croix"">Saint Croix</option>" & strLE & _
			"<option value=""Saint Helena"">Saint Helena</option>" & strLE & _
			"<option value=""Saint Kitts and Nevis"">Saint Kitts and Nevis</option>" & strLE & _
			"<option value=""Saint Lucia"">Saint Lucia</option>" & strLE & _
			"<option value=""Saint Pierre and Miquelon"">Saint Pierre and Miquelon</option>" & strLE & _
			"<option value=""Saint Vincent and Grenadi"">Saint Vincent and Grenadi</option>" & strLE & _
			"<option value=""San Marino"">San Marino</option>" & strLE & _
			"<option value=""Sao Tome and Principe"">Sao Tome and Principe</option>" & strLE & _
			"<option value=""Saudi Arabia"">Saudi Arabia</option>" & strLE & _
			"<option value=""Senegal"">Senegal</option>" & strLE & _
			"<option value=""Seychelles"">Seychelles</option>" & strLE & _
			"<option value=""Sierra Leone"">Sierra Leone</option>" & strLE & _
			"<option value=""Singapore"">Singapore</option>" & strLE & _
			"<option value=""Slovakia"">Slovakia</option>" & strLE & _
			"<option value=""Slovenia"">Slovenia</option>" & strLE & _
			"<option value=""Solomon Islands"">Solomon Islands</option>" & strLE & _
			"<option value=""Somalia Northern Region"">Somalia Northern Region</option>" & strLE & _
			"<option value=""Somalia Southern Region"">Somalia Southern Region</option>" & strLE & _
			"<option value=""South Africa"">South Africa</option>" & strLE & _
			"<option value=""South Sandwich Islands"">South Sandwich Islands</option>" & strLE & _
			"<option value=""Spain"">Spain</option>" & strLE & _
			"<option value=""Sri Lanka"">Sri Lanka</option>" & strLE & _
			"<option value=""Sudan"">Sudan</option>" & strLE & _
			"<option value=""Suriname"">Suriname</option>" & strLE & _
			"<option value=""Swaziland"">Swaziland</option>" & strLE & _
			"<option value=""Sweden"">Sweden</option>" & strLE & _
			"<option value=""Switzerland"">Switzerland</option>" & strLE & _
			"<option value=""Syria"">Syria</option>" & strLE & _
			"<option value=""Taiwan"">Taiwan</option>" & strLE & _
			"<option value=""Tajikistan"">Tajikistan</option>" & strLE & _
			"<option value=""Tanzania"">Tanzania</option>" & strLE & _
			"<option value=""Thailand"">Thailand</option>" & strLE & _
			"<option value=""Togo"">Togo</option>" & strLE & _
			"<option value=""Tonga"">Tonga</option>" & strLE & _
			"<option value=""Trinidad and Tobago"">Trinidad and Tobago</option>" & strLE & _
			"<option value=""Tunisia"">Tunisia</option>" & strLE & _
			"<option value=""Turkey"">Turkey</option>" & strLE & _
			"<option value=""Turkmenistan"">Turkmenistan</option>" & strLE & _
			"<option value=""Turks and Caicos Islnd"">Turks and Caicos Islnd</option>" & strLE & _
			"<option value=""Tuvalu"">Tuvalu</option>" & strLE & _
			"<option value=""USA"">USA</option>" & strLE & _
			"<option value=""Uganda"">Uganda</option>" & strLE & _
			"<option value=""Ukraine"">Ukraine</option>" & strLE & _
			"<option value=""United Arab Emirates"">United Arab Emirates</option>" & strLE & _
			"<option value=""United Kingdom"">United Kingdom</option>" & strLE & _
			"<option value=""Uruguay"">Uruguay</option>" & strLE & _
			"<option value=""Uzbekistan"">Uzbekistan</option>" & strLE & _
			"<option value=""Vanuatu"">Vanuatu</option>" & strLE & _
			"<option value=""Vatican City"">Vatican City</option>" & strLE & _
			"<option value=""Venezuela"">Venezuela</option>" & strLE & _
			"<option value=""Vietnam"">Vietnam</option>" & strLE & _
			"<option value=""Virgin Islands (United Kingdom)"">Virgin Islands (United Kingdom)</option>" & strLE & _
			"<option value=""Wallis and Futuna Islands"">Wallis and Futuna Islands</option>" & strLE & _
			"<option value=""Western Sahara"">Western Sahara</option>" & strLE & _
			"<option value=""Western Samoa"">Western Samoa</option>" & strLE & _
			"<option value=""Yemen"">Yemen</option>" & strLE & _
			"<option value=""Yugoslavia"">Yugoslavia</option>" & strLE & _
			"<option value=""Zambia"">Zambia</option>" & strLE & _
			"<option value=""Zimbabwe"">Zimbabwe</option>" & strLE
%>
