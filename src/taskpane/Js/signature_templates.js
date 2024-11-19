// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function get_template_A_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str +='<table border="0" cellpadding="5" cellspacing="5"><tbody><tr><td valign="top"><font size="3" color="#17365d" face="Arial">';
  str +='<strong>'+ user_info.name +'</strong></font>';
  str +='<br><font size="2" face="Arial">'+ user_info.job +'</font><br>';
  str +='<font size="3" color="#17365d" face="Arial">';
  str += is_valid_data(user_info.pronoun) ? "<strong>" + user_info.pronoun : "";
  str += '</strong></font><br><font size="2" face="Arial">Tel.:';
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += user_info.email;
  str +='</font></td></tr><tr><td><table border="0" cellpadding="0" cellspacing="0"><tbody><tr><td width="240" height="81"><img src="https://laramirez0310.github.io/OfficeSignature/assets/marca-pucmm.jpg" width="auto" height="70" alt="Pontificia Universidad Católica Madre y Maestra"></td><td width="15"></td><td style="padding:0 0 0 15px;border-left-style:solid;border-left-width:1pt;border-left-color:#7f7f7f"><p><font size="2" face="Arial"><strong>Campus de Santiago:</strong><br>Autopista Duarte km. 1½, Santiago, R.D.';
  str +='<br><br><strong>Campus de Santo Domingo:</strong><br>Av. Abraham Lincoln. esq. Av. Bolívar, Santo Domingo, R.D.</font></p></td></tr></tbody></table></td></tr><tr><td height="70" align="left" valign="middle"><table border="0" cellpadding="0" cellspacing="0"><tbody><tr><td width="600" height="70"><img src="https://pucmm.edu.do/PublishingImages/firma-60aniv/BannerFirma60aniv.png" width="600" height="70" alt="60 Aniversario PUCMM"></td><td width="15"></td><td class="social" style="display: flex; align-items: center;justify-content: space-around;" width="150" height="70"><a class="social-icons" href="http://www.facebook.com/pucmm/" target="_blank"><img style="margin:2px;" src="https://www.pucmm.edu.do/PublishingImages/firma-60aniv/facebook.png" alt="Facebook PUCMM" width="24" height="24"></a><a class="social-icons" href="http://twitter.com/pucmm/" target="_blank"><img src="https://www.pucmm.edu.do/PublishingImages/firma-60aniv/twitter.png" style="margin:2px;" alt="Twitter PUCMM" width="24" height="25"></a><a class="social-icons" href="http://www.instagram.com/pucmm" target="_blank"><img src="https://www.pucmm.edu.do/PublishingImages/firma-60aniv/instagram.png" style="margin:2px;" alt="Instagram PUCMM" width="24" height="25"></a><a class="social-icons" href="http://www.youtube.com/pucmmtv/" target="_blank"><img src="https://www.pucmm.edu.do/PublishingImages/firma-60aniv/youtube.png" style="margin:2px;" alt="Youtube PUCMM" width="24" height="25"></a><a class="social-icons" href="https://www.linkedin.com/edu/school?id=12020" target="_blank"><img src="https://www.pucmm.edu.do/PublishingImages/firma-60aniv/linkedin.png" style="margin:2px;" alt="Linkedin PUCMM" width="24" height="25"></a></td></tr></tbody></table></td></tr><tr><td><img src="https://www.pucmm.edu.do/PublishingImages/firma-60aniv/Green.gif" width="14" height="14">';
  str +='<font color="#7F7F7F" size="3" face="Arial">No me imprimas si no es necesario.</font></td></tr><tr><td><p style="margin:0"><font color="#7d7d7d" size="1" face="Arial">NOTA DE CONFIDENCIALIDAD: La información transmitida, incluidos los archivos adjuntos, está dirigida solo a la persona o entidad a la que ha sido remitida y puede contener información confidencial y/o privilegiada. Cualquier difusión u otro uso de la misma, o tomar cualquier acción basada en esta información por personas o entidades distintas al destinatario, está prohibido. Si ha recibido este mensaje por error, por favor contactar al remitente y destruya cualquier copia de esta información.<br>CONFIDENTIALITY NOTE: The information transmitted, including attachments, is intended only for the person or entity to which it is addressed and may contain confidential and/or privileged material. Any review, retransmission, dissemination or other use of, or taking of any action in reliance upon this information by persons or entities other than the intended recipient is prohibited. If you received this in error, please contact the sender and destroy any copies of this information.</font></p></td></tr></tbody></table>';


  return str;
}

function get_template_B_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += "<table style='display:none;'>";
  str +=   "<tr>";
  str +=     "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://laramirez0310.github.io/OfficeSignature/assets/marca-pucmm.jpg' alt='Logo' /></td>";
  str +=     "<td style='padding-left: 5px;'>";
  str +=	   "<strong>" + user_info.name + "</strong>";
  str +=     is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str +=     "<br/>";
  str +=	   user_info.email + "<br/>";
  str +=	   is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str +=     "</td>";
  str +=   "</tr>";
  str += "</table>";

  return str;
}

function get_template_C_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += user_info.name;
  
  return str;
}
