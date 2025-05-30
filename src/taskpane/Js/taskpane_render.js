// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

let _display_name;
let _job_title;
let _phone_number;
let _email_id;
let _greeting_text;
let _preferred_pronoun;
let _message;

Office.initialize = function(reason)
{
  on_initialization_complete();
}

function on_initialization_complete()
{
	$(document).ready
	(
		function()
		{
      _output = $("textarea#output");
      _display_name = $("input#display_name");
      _email_id = $("input#email_id");
      _job_title = $("input#job_title");
      _phone_number = $("input#phone_number");
      _greeting_text = $("input#greeting_text");
      _preferred_pronoun = $("input#preferred_pronoun");
      _message = $("p#message");

      prepopulate_from_userprofile();
      cargar_datos();
		}
	);
}

function prepopulate_from_userprofile()
{
  _display_name.val(Office.context.mailbox.userProfile.displayName);
  _email_id.val(Office.context.mailbox.userProfile.emailAddress);
  /*_job_title.val(Office.context.mailbox.userProfile.jobTitle);
  _phone_number.val(Office.context.MailboxEnums.EntityType.PhoneNumber);*/
 
}
var item = Office.context.mailbox.item;
// Get an array of strings that represent contacts in the current item's body.
var contacts = item.getEntitiesByType(Office.MailboxEnums.EntityType.Contact);
console.log("There are " + contacts.length + " contacts.")
contacts.forEach(function (contact) {
    console.log("Person name: " + JSON.stringify(contact.personName));
    console.log("Business name: " + JSON.stringify(contact.businessName));
    console.log("Addresses: " + JSON.stringify(contact.addresses));
    console.log("Phone numbers: " + JSON.stringify(contact.phoneNumbers));
    console.log("Email addresses: " + JSON.stringify(contact.emailAddresses));
    console.log("Urls: " + JSON.stringify(contact.urls));
});

function load_saved_user_info()
{
  let user_info_str = localStorage.getItem('user_info');
  if (!user_info_str)
  {
    user_info_str = Office.context.roamingSettings.get('user_info');
  }

  if (user_info_str)
  {
    const user_info = JSON.parse(user_info_str);

    _display_name.val(user_info.name);
    _email_id.val(user_info.email);
    _job_title.val(user_info.job);
    _phone_number.val(user_info.phone);
    _greeting_text.val(user_info.greeting);
    _preferred_pronoun.val(user_info.pronoun);
    
  }
}

function display_message(msg)
{
  _message.text(msg);
}

function clear_message()
{
  _message.text("");
}

function is_not_valid_text(text)
{
  return text.length <= 0;
}

function is_not_valid_email_address(email_address)
{
  let email_address_regex = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/;
  return is_not_valid_text(email_address) || !(email_address_regex.test(email_address));
}

function form_has_valid_data(name, email)
{
  if (is_not_valid_text(name))
  {
    display_message("Por favor ingrese un nombre valido.");
    return false;
  }

  if (is_not_valid_email_address(email))
  {
    display_message("Por favor, introduce una dirección de correo válida.");
    return false;
  }

  return true;
}
 
function navigate_to_taskpane_assignsignature()
{
  window.location.href = 'assignsignature.html';
}

function create_user_info()
{
  let name = _display_name.val().trim();
  let email = _email_id.val().trim();

  clear_message();

  if (form_has_valid_data(name, email))
  {
    clear_message();
 
    let user_info = {};

    user_info.name = name;
    user_info.email = email;
    user_info.job =  _job_title.val().trim();
    user_info.phone = _phone_number.val().trim();
    //user_info.greeting = _greeting_text.val().trim();
    user_info.pronoun = _preferred_pronoun.val().trim();

    if (user_info.pronoun !== "")
    {
      user_info.pronoun = "" + user_info.pronoun + "";
    }

    
    localStorage.setItem('user_info', JSON.stringify(user_info));
    save_signature_settings();
  }
}

function clear_all_fields()
{
  _display_name.val("");
  _email_id.val("");
  _job_title.val("");
  _phone_number.val("");
  _greeting_text.val("");
  _preferred_pronoun.val("");
}

function clear_all_localstorage_data()
{
  localStorage.removeItem('user_info');
  localStorage.removeItem('newMail');
  localStorage.removeItem('reply');
  localStorage.removeItem('forward');
  localStorage.removeItem('override_olk_signature');
}

function clear_roaming_settings()
{
  Office.context.roamingSettings.remove('user_info');
  Office.context.roamingSettings.remove('newMail');
  Office.context.roamingSettings.remove('reply');
  Office.context.roamingSettings.remove('forward');
  Office.context.roamingSettings.remove('override_olk_signature');

  Office.context.roamingSettings.saveAsync
  (
    function (asyncResult)
    {
      //console.log("clear_roaming_settings - " + JSON.stringify(asyncResult));

      let message = "¡Todas las configuraciones se restablecieron con éxito! Este complemento no insertará ninguna firma. Puede cerrar este panel ahora.";
      if (asyncResult.status === Office.AsyncResultStatus.Failed)
      {
        message = "No se pudo restablecer. Inténtalo de nuevo.";
      }

      display_message(message);
    }
  );
}

function reset_all_configuration()
{
  clear_all_fields();
  clear_all_localstorage_data();
  clear_roaming_settings();
}