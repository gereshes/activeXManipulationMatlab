function emailGeneral(subject,recipientEmailAdressBundle,body )
%This function is a general porpoise formatter that opens outlook on your
%desktop and sends out an email from it.
%
%subject,recipientEmailAdress,body - Are all strings
%
% Ari Rubinsztejn
% a.rubin1225@gmail.com
% www.gereshes.com


try
    %Checks if there is already a running activeX session running
    outlookHandle=actxGetRunningServer('Outlook.Application');
catch
    %Starts an activeX session if there isnt one running
    outlookHandle=actxserver('Outlook.Application');
end
numOfEmails=length(recipientEmailAdressBundle);
for emails=1:numOfEmails
    recipientEmailAdress=recipientEmailAdressBundle{emails};
    %Fills in the email and sends it
    mail=outlookHandle.CreateItem('olMail');
    mail.Subject=subject;
    mail.To=recipientEmailAdress;
    mail.BodyFormat = 'olFormatHTML';
    mail.HTMLBody= body;
    mail.Send;
end


end

