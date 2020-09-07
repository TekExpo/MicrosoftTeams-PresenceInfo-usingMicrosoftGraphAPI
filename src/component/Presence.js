import React, { Component } from "react";
import { NormalPeoplePicker } from "office-ui-fabric-react/lib/Pickers";
import { Persona, PersonaPresence } from "office-ui-fabric-react/lib/Persona";
import { Button } from "office-ui-fabric-react/lib/Button";
import { Label } from "office-ui-fabric-react/lib/Label";
import { Spinner } from "office-ui-fabric-react/lib/Spinner";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import { DetailsList } from "office-ui-fabric-react/lib/DetailsList";
import {
  MessageBar,
  MessageBarType
} from "office-ui-fabric-react/lib/MessageBar";

export default class PresenceExample extends Component {
  constructor() {
    super();

    // Set the initial state for the picker data source.
    // The people list is populated in the _onFilterChanged function.
    this._peopleList = null;
    this._searchResults = [];
    this._lstitems = [];

    // Helper that uses the JavaScript SDK to communicate with Microsoft Graph.
    this.sdkHelper = window.sdkHelper;

    this._showError = this._showError.bind(this);
    this.state = {
      selectedPeople: [],
      isLoadingPeople: true,
      isLoadingPics: true
    };  
    
  }

  // Map user properties to persona properties.
  _mapUsersToPersonas(users, useMailProp) {
    return users.map(p => {
      // The email property is returned differently from the /users and /people endpoints.
      let email = useMailProp ? p.mail : p.emailAddresses[0].address;
      let persona = new Persona();

      persona.primaryText = p.displayName;
      persona.secondaryText = email || p.userPrincipalName;
      persona.presence = PersonaPresence.none; // Presence isn't supported in Microsoft Graph yet
      persona.imageInitials =
        !!p.givenName && !!p.surname
          ? p.givenName.substring(0, 1) + p.surname.substring(0, 1)
          : p.displayName.substring(0, 1);
      persona.initialsColor = Math.floor(Math.random() * 15) + 0;
      persona.props = { id: p.id };

      return persona;
    });
  }

  // Gets the profile photo for each user.
  _getPics(personas) {
    // Make suggestions available before retrieving profile pics.
    this.setState({
      isLoadingPeople: false
    });

    this.sdkHelper.getProfilePics(personas, err => {
      this.setState({
        isLoadingPics: false
      });
    });
  }

  // Build and send the email to the selected people.
  _getPresenceOfSelectedPeople() {
    const recipients = this.state.selectedPeople.map(r => r.props.id);
    const recpName = this.state.selectedPeople[0].primaryText;
    this.sdkHelper.getPresence(recipients, (err, toRecipients) => {
      this._processItems(err, toRecipients, recpName);
    });
  }

  // Map properties metadata to list items.
  _processItems(err, res, recp) {
    if (!err) {
      const response = res.value;
      let nextLink = null;
      const listItems = response.map(f => {
        return {
          Id: f.id,
          Status: f.activity,
          Availability: f.availability,
          DisplayName: recp
        
        };
      });
      this._lstitems = this._lstitems.concat(listItems);
      this.setState({
        listItems: this._lstitems,
        isLoading: !!nextLink,
        nextPageToken: nextLink
      });
    } else this._showError(err);
  }

  // Handler for when text is entered into the picker control.
  // Populate the people list.
  _onFilterChanged(filterText, items) {
    
    if (this._peopleList) {
      return filterText? this._peopleList
            .concat(this._searchResults)
            .filter(
              item =>
                item.primaryText
                  .toLowerCase()
                  .indexOf(filterText.toLowerCase()) === 0
            )
            .filter(item => !this._listContainsPersona(item, items))
        : [];
    } else {
      return new Promise((resolve, reject) =>
        this.sdkHelper.getPeople((err, people) => {
          if (!err) {
            this._peopleList = this._mapUsersToPersonas(people, false);
            this._getPics(this._peopleList);
            resolve(this._peopleList);
          } else {
            // this._showError(err);
          }
        })
      ).then(value =>
        value
          .concat(this._searchResults)
          .filter(
            item =>
              item.primaryText
                .toLowerCase()
                .indexOf(filterText.toLowerCase()) === 0
          )
          .filter(item => !this._listContainsPersona(item, items))
      );
    }
  }

  // Remove currently selected people from the suggestions list.
  _listContainsPersona(persona, items) {
    if (!items || !items.length || items.length === 0) {
      return false;
    }
    return (
      items.filter(item => item.primaryText === persona.primaryText).length > 0
    );
  }

  // Handler for when the Search button is clicked.
  // This sample returns the first 20 matches as suggestions.
  _onGetMoreResults(searchText) {
    this.setState({
      isLoadingPeople: true,
      isLoadingPics: true
    });
    return new Promise(resolve => {
      this.sdkHelper.searchForPeople(
        searchText.toLowerCase(),
        (err, people) => {
          if (!err) {
            this._searchResults = this._mapUsersToPersonas(people, true);
            this.setState({
              isLoadingPeople: false
            });
            this._getPics(this._searchResults);
            resolve(this._searchResults);
          }
        }
      );
    });
  }

  // Handler for when the selection changes in the picker control.
  // This sample updates the list of selected people and clears any messages.
  _onSelectionChanged(items) {
    this.setState({
      result: null,
      selectedPeople: items
    });
  }

  getTextStyle(status)
  {
    if(status==="Away")
     {
      const styles = {
        color: "gray",
       fontWeight: "bold",
        fontSize: "15px"
    };
     return styles;
      }
      else if(status==="Offline") {
        const styles = {
          color: "blue",
         
          fontWeight: "bold",
          fontSize: "15px"
      };
       return styles;
       }
       else if(status==="Available") {
        const styles = {
          color: "green",
         
          fontWeight: "bold",
          fontSize: "15px"
      };
       return styles;
       }
       else if(status==="Busy") {
        const styles = {
          color: "red",
         fontWeight: "bold",
          fontSize: "15px"
      };
       return styles;
       }
       else if(status==="DoNotDisturb") {
        const styles = {
          color: "brown",
         fontWeight: "bold",
          fontSize: "15px"
      };
       return styles;
       }
       else{
        const styles = {
          color: "#F0A202",
         fontWeight: "bold",
          fontSize: "15px"
      };
       return styles;
       }
     }
 

  // Renders the people picker using the NormalPeoplePicker template.
  render() {
    return (
      <div>
        <h3>User Presence example</h3>

        <Label>
          <b>Search</b>.
        </Label>

        <NormalPeoplePicker
          onResolveSuggestions={this._onFilterChanged.bind(this)}
          pickerSuggestionsProps={{
            suggestionsHeaderText: "Suggested People",
            noResultsFoundText: "No results found",
            searchForMoreText: "Search",
            loadingText: "Loading...",
            isLoading: this.state.isLoadingPics
          }}
          getTextFromItem={persona => persona.primaryText}
          onChange={this._onSelectionChanged.bind(this)}
          onGetMoreResults={this._onGetMoreResults.bind(this)}
          className="ms-PeoplePicker"
          key="normal-people-picker"
        />
        <br />
        <Button
          buttonType={0}
          onClick={this._getPresenceOfSelectedPeople.bind(this)}
          disabled={!this.state.selectedPeople.length > 0}
        >
          Get Presence
        </Button>
             

        <div><ul>{this._lstitems.map(listItem => <div> <li class="ms-DetailsHeader-cell is-actionable">{listItem.DisplayName} - Status : <span style={this.getTextStyle(listItem.Status)}>{listItem.Status}</span></li></div>)}</ul></div>
        <br />
        {this.state.result && (
          <MessageBar messageBarType={this.state.result.type}>
            {this.state.result.text}
          </MessageBar>
        )}
      </div>
    );
  }

  // Show the results of the `/me/people` query.
  // For sample purposes only.
  _showPeopleResults() {
    let message = "Query loading. Please try again.";
    if (!this.state.isLoadingPeople) {
      const people = this._peopleList.map(p => {
        return `\n${p.primaryText}`;
      });
      message = people.toString();
    }
    alert(message);
  }

  // Configure the error message.
  _showError(err) {
    this.setState({
      result: {
        type: MessageBarType.error,
        text: `Error ${err.statusCode}: ${err.code} - ${err.message}`
      }
    });
  }

  

}
