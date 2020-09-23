<div align="center">

## Handy Form Utility


</div>

### Description

This helps you verify all kinds of forms using regular expressions.
 
### More Info
 
'* EMAIL		=	xyxyxy.xyxyxy@xyxy.xyx

'* POSTAL		=	yyyy+

'* NAME			=	xxxxxx xxxxx

'* STREET		=	xxxxxxx yy

'* COUNTRY		=	(all countries in ENGLISH only)

'* ALPHASTRING	=	xxxxx x xx xx xx x xxx

It simply gives you back TRUE OR FALSE for the user's input using regular expressions. But it's strict and.. well pretty handy stuff.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Shrike](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shrike.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Validation/ Processing](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/validation-processing__4-16.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/shrike-handy-form-utility__4-6560/archive/master.zip)





### Source Code

```
	FUNCTION frmCheck(strType,strCheck)
					DIM objRegExp
					SET objRegExp			=	NEW RegExp
					objRegExp.IgnoreCase	=	TRUE
					SELECT CASE strType
						CASE "email"		objRegExp.Pattern	=	"^([A-Za-z0-9_.]{3,15})\@(\w{3,15})\.(\w{3,15})$"
						CASE "postal"		objRegExp.Pattern	=	"^([0-9]{4,})$"
						CASE "name"			objRegExp.Pattern	=	"^([A-Za-z]{3,15})(\s)([A-Za-z]{3,15})$"
						CASE "street"		objRegExp.Pattern	=	"^([A-Za-z]{3,15})(\s)([A-Za-z0-9\/\.]{1,8})$"
						CASE "country"		objRegExp.Pattern	=	"^(Afghanistan|Albania|Algeria|Andorra|Angola|Anguilla|Antarctica|Antigua|Argentina|Armenia|Aruba|Ashmore|Australia|Austria|Azerbaijan|Bahamas|Bahrain|Baker|Baker Island|Barbados|Bassas|Bangladesh|Belarus|Belgium|Belize|Benin|Bermuda|Bhutan|Botswana|Bolivia|Bosnia|Bouvet|Bouvet Island|Brazil|British|Brunei|Bulgaria|Burkina|Burkina Faso|Burma|Burundi|Cambodia|Cameroon|Canada|Capeverde|Cape Verde|Cayman|Cayman Islands|Centralaf|Central African Republic|Chad|Chile|China|Christmas|Christmas Island|Cocos|Colombia|Comoros|Congo|Congodem|Cook|Costarica|Costa Rica|Cote|Cote D`Ivoire|Ivory Coast|Croatia|Cuba|Cyprus|Czechrep|Czech Republic|Denmark|Djibouti|Dominica|Dominican|Dominican Republic|Ecuador|Egypt|El Salvador|Equatorial Guinea|Eritrea|Estonia|Ethiopia|Falkland|Falkland Islands|Faroe|Faroe Islands|Fiji|Finland|France|French Guiana|French Polynesia|French Southern Territories|Fyrom|F.Y.R.O.M.|Gabon|Gambia|Gaza|Gaza Strip|Georgia|Germany|Ghana|Gibraltar|UK|Great Britain|Greece|Greenland|Grenada|Guadeloupe|Guam|Guatemala|Guernsey|Guinea|Guinea-Bissau|Guyana|Haiti|Heard|Heard & Mcdonald Islands|Honduras|Hongkong|Hong Kong|Hungary|Iceland|India|Indonesia|Iran|Iraq|Ireland|Israel|Italy|Jamaica|Japan|Jordan|Kazakhstan|Kenya|Kiribati|Korea|Korea South|Korea North|Kuwait|Kyrgyzstan|Laos|Latvia|Lebanon|Lesotho|Liechten|Liechtenstein|Liberia|Libya|Lithuania|Luxembourg|Macau|Madagascar|Malawi|Malaysia|Maldives|Mali|Malta|Marshall|Marshall Islands|Martinique|Mauritania|Mauritius|Mayotte|Mexico|Micronesia|Monaco|Moldova|Morocco|Mongolia|Montserrat|Mozambique|Myanmar|Namibia|Nauru|Nepal|Netherlands|Antilles|Netherlands Antilles|Caledonia|New Caledonia|Nicaragua|Niger|Nigeria|Niue|Norfolk|Norfolk Island|Mariana|Northern Mariana Islands|Norway|Oman|Pakistan|Palau|Palestine|Panama|Papua|Papua New Guinea|Paraguay|Peru|Philippines|Pitcairn|Pitcairn Islands|Poland|Portugal|Puertorico|Puerto Rico|Qatar|Reunion|Romania|Russian|Russian Federation|Rwanda|Kitts|Saint Kitts & Nevis|Lucia|Saint Lucia|Vincent|Saint Vincent |Marino|San Marino|Tome|Sao Tome & Principe|Saudi|Saudi Arabia|Senegal|Seychelles|Sierra|Sierra Leone|Singapore|Slovenia|Slovak|Slovak Republic|Solomon|Solomon Islands|Somalia|South Africa|Spain|Srilanka|Sri Lanka|Sthelena|St. Helena|Stpierre|St. Pierre & Miquelon|Sudan|Suriname|Svalbard|Svalbard & Jan Mayen Isl.|Swaziland|Sweden|Switzerland|Syria|Taiwan|Tajikistan|Tanzania|Thailand|Togo|Tokelau|Tonga|Trinidad|Trinidad & Tobago|Tunisia|Turkey|Turkmenistan|Turks|Tuvalu|Uganda|Ukraine|United Arab Emirates|UK|United Kingdom|USA|United States|Uruguay|Uzbekistan|Vanuatu|Vatican|Venezuela|Vietnam|Virgin Islands|Wallis|Western Sahara|Westbank|West Bank|Yemen|Yugoslavia|Zambia|Zimbabwe)$"
						CASE "alphastring"	objRegExp.Pattern	=	"^([A-Za-z\s]{1,255})$"
						CASE ELSE
							 frmCheck	=	"False (invalid strType)"
							 EXIT FUNCTION
					END SELECT
					frmCheck			=	objRegExp.Test(cStr(strCheck))
	END FUNCTION
	Use :
	IF frmCheck("email",strEmail) = TRUE AND frmCheck("country",strCountry) = TRUE THEN ...
	Try it simple with :
	Response.Write frmCheck("email","blabla@bla.com")
	Response.Write frmCheck("country","zimbabwe")
	Response.Write frmCheck("name","jennifer lopez")
	Response.Write frmCheck("street","aspstreet 17a")
```

