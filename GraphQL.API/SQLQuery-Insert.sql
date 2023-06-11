insert into  TechEventInfo( 
    EventName ,    
    Speaker ,    
    EventDate
)
select 'Sitecore Sugcon-India', 'Mike Duglus', getdate() union
select 'Azure Open AI Integration-Denmark', 'Stefan Praise', getdate() union
select 'Cloud Native development-London', 'Ryna Jacob', getdate() 

insert into  Participant(ParticipantName,Email,Phone)
select  'Rajeev Singh', 'rajeev.singh@gmail.com', '+91 6780004561' union
select  'Robin Muffet', 'robin.muffet@gmail.com', '+01 789004521' union
select  'Sanjay Gupta', 'sanjay.gupta@gmail.com', '+91 9778004561' union
select  'Marry White', 'w.marry@gmail.com', '+21 589004521' union
select  'Abdul Kadir', 'abdul.kadir91@gmail.com', '+91 8780004561' union
select  'Rahul Agrwal', 'rahul.agr@gmail.com', '+91 9689004521' 

insert into  EventParticipants  
(     
    EventId,  
    ParticipantId
)
select 1,1 union
select 1,2 union
select 2,3 union
select 2,4 union
select 3,5 union
select 3,6 