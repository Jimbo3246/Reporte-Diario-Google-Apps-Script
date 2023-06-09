function sendFileToWhatsAppGroup(fileUrl, messageText) {
 
  var twilioAccountSid = "AC262f923cdced238c3b36e8d041790fcf";
  var twilioAuthToken = "[AuthToken]";
  var twilioSandboxNumber = "your_twilio_sandbox_number";
  var whatsAppGroupNumber = "whatsapp:your_whatsapp_group_number";

  var twilioClient = Twilio(twilioAccountSid, twilioAuthToken);

  var message = {
    body: messageText,
    from: 'whatsapp:' + twilioSandboxNumber,
    mediaUrl: fileUrl
  };

  twilioClient.messages.create({
    to: whatsAppGroupNumber,
    from: 'whatsapp:' + twilioSandboxNumber,
    mediaUrl: fileUrl,
    body: messageText
  });
}