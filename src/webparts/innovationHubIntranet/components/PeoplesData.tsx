import * as React from 'react';
import { Label } from '@fluentui/react/lib/Label';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { IBasePickerSuggestionsProps , NormalPeoplePicker, ValidationState } from 'office-ui-fabric-react/lib/Pickers';

let peoplesnames = [
    {
      key: 1,
      imageUrl:
        "/_layouts/15/userphoto.aspx?size=S&accountname=kalimuthu@chandrudemo.onmicrosoft.com",
      text: "kali muthu",
      ID:0,
      secondaryText: "Designer",
      isValid: true
    },
    {
      key: 2,
      imageUrl:
        "/_layouts/15/userphoto.aspx?size=S&accountname=chandru@chandrudemo.onmicrosoft.com",
      text: "chandra moorthy",
      secondaryText: "Designer",
      ID:0,
      isValid: true
    },
  ];

function PeoplesData(props): React.ReactElement<[]>
{
    const [peopleList, setPeopleList] = React.useState<IPersonaProps[]|any>(peoplesnames);

    React.useEffect(()=>
    {
        //getGroupFromList();

    },[]);

    function GetUserDetails(filterText)
    {
        

        var result = peopleList.filter((value, index, self) =>
        index === self.findIndex((t) => (
          t.ID === value.ID
        ))
      )

      return result.filter(item => doesTextStartWith(item.text as string, filterText));
    }

    function doesTextStartWith(text: string, filterText: string): boolean {
        return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
      }

    return(<div>
        <Label>Add User</Label>
        <NormalPeoplePicker onResolveSuggestions={GetUserDetails}
        /*createGenericItem={()=>{
            return(<div>dshsdsdhsdhshsdhsdh</div>)
        }}
        onValidateInput={()=>{
            return ValidationState.valid;
        }}*/
        itemLimit={1}
        onChange={(items)=>{
            console.log(items);
            //props.update(items);
        }}
        /*onRemoveSuggestion={(items)=>{
            console.log(items);
        }}
       onItemSelected={(items)=>{//which is used to get the selected item
            console.log(items);
            return items;
       }}*/
       // defaultSelectedItems={peopleList.slice(0,1)}//which is used for selected items
        inputProps={{
            onBlur: (ev: React.FocusEvent<HTMLInputElement>) =>{ },
            onFocus: (ev: React.FocusEvent<HTMLInputElement>) =>{},
            'aria-label': 'People Picker',
          }}
        />
    </div>)


    async function getGroupFromList()
    {
        var groups=[];
        var users=[];
        await props.spcontext.lists.getByTitle("ConfigUsers")
        .items.select("GroupName/ID","Category").filter("Category eq 'WF'").expand("GroupName").get().then(async function(data)
        {
           if(data.length>0)
           {
                await data.forEach(async item => 
                {
                    if(item.GroupName.length>0)
                    {
                        await item.GroupName.forEach(async element => 
                        {
                            await groups.push(element.ID);    
                        });

                        await groups.forEach(async groupid => 
                        {
                            await props.spcontext.siteGroups.getById(groupid).users.get().then(async function(result) 
                                {
                                    
                                    
                                    for (var i = 0; i < result.length; i++) 
                                    {
                                        
                                        var userdetails={
                                            key: 1,
                                            imageUrl:
                                              "/_layouts/15/userphoto.aspx?size=S&accountname="+result[i].Email,
                                            text: result[i].Title,
                                            secondaryText: result[i].Email,
                                            ID:result[i].Id,
                                            isValid: true
                                          }
                                          
                                          await users.push(userdetails);
                                    }
                                    
                                }).catch(function(err) 
                                {
                                    //alert("Group not found: " + err);
                                    console.log("Group not found: " + err);
                                });
                        });
                    }
                });
        
                setPeopleList(users);
           }
           else
           {
                setPeopleList([]);
           }

           
            
        }).catch(function(error){
            alert(error)
        })
    }

}

export default PeoplesData;