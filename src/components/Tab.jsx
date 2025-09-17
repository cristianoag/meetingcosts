import React from "react";
import { app, teamsCore } from "@microsoft/teams-js";
import MediaQuery from "react-responsive";
import "./App.css";

class Tab extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      context: {},
      participants: [],
      participantsLoading: false,
      participantsError: null,
      meetingIdDecoded: "",
      rawResponse: "",
      totalCost: 0,
    };
  }

  //React lifecycle method that gets called once a component has finished mounting
  //Learn more: https://reactjs.org/docs/react-component.html#componentdidmount
  componentDidMount() {
    app.initialize().then(() => {
      // Notifies that the app initialization is successfully and is ready for user interaction.
      app.notifySuccess();

      // Get the user context from Teams and set it in the state
      app.getContext().then(async (context) => {
        const meetingId = context.meeting.id;
        let meetingId_decoded = "";
        if (meetingId) {
          try {
            meetingId_decoded = atob(meetingId).replace(/^0#|#0$/g, "");
          } catch (e) {
            console.warn("Failed to decode meeting id", e);
          }
        }
        this.setState(
          {
            meetingId: meetingId,
            meetingIdDecoded: meetingId_decoded,
            userName: context.user.userPrincipalName,
          },
          () => {
            // After state set, fetch participants
            this.getParticipants();
          }
        );

        // Enable app caching.
        // App Caching was configured in this sample to reduce the reload time of your app in a meeting.
        // To learn about limitations and available scopes, please check https://learn.microsoft.com/en-us/microsoftteams/platform/apps-in-teams-meetings/app-caching-for-your-tab-app.
        if (context.page.frameContext === "sidePanel") {
          teamsCore.registerOnLoadHandler((context) => {
            // Use context.contentUrl to route to the correct page.
            app.notifySuccess();
          });

          teamsCore.registerBeforeUnloadHandler((readyToUnload) => {
            // Dispose resources here if necessary.
            // Notify readiness by invoking readyToUnload.
            readyToUnload();
            return true;
          });
        }
      });
    });
    // Next steps: Error handling using the error object
  }

  getParticipants = async () => {
    const { meetingIdDecoded } = this.state;
    if (!meetingIdDecoded) return;
    this.setState({ participantsLoading: true, participantsError: null });
    try {
      const logicAppUrl = process.env.REACT_APP_AAD_LOGIC_APP_URL;
      if (!logicAppUrl) {
        throw new Error("Environment variable AAD_LOGIC_APP_URL is not defined.");
      }
      const response = await fetch(
        logicAppUrl,
        {
          method: "POST",
          headers: {
        "Content-Type": "application/json",
        "Accept": "application/json, text/plain, */*",
          },
          body: JSON.stringify({
        meetingId: meetingIdDecoded,
          }),
        }
      );

      if (!response.ok) {
        const text = await response.text().catch(() => "");
        throw new Error(`HTTP ${response.status} ${response.statusText} - ${text.substring(0,200)}`);
      }
      const raw = await response.text();
      this.setState({ rawResponse: raw });
      if (!raw) {
        this.setState({ participants: [], participantsLoading: false });
        return;
      }
      let data;
      try {
        data = JSON.parse(raw);
      } catch (parseErr) {
        // If parsing top-level fails, treat raw as participants JSON array string
        console.warn('Top-level JSON parse failed, attempting direct participants parse', parseErr);
        try {
          const possibleArray = JSON.parse(raw);
          if (Array.isArray(possibleArray)) {
            this.setState({ participants: possibleArray, participantsLoading: false });
            return;
          }
        } catch (_) {}
        this.setState({ participantsError: 'Failed to parse response JSON', participantsLoading: false });
        return;
      }
      let participantsArray = [];
      if (Array.isArray(data)) {
        participantsArray = data;
      } else if (data && Array.isArray(data.participants)) {
        participantsArray = data.participants;
      } else if (data && typeof data.participants === 'string') {
        // Attempt to parse stringified JSON array
        try {
          const parsed = JSON.parse(data.participants);
          if (Array.isArray(parsed)) {
            participantsArray = parsed;
          }
        } catch (e) {
          console.warn('Failed to parse participants string', e);
          this.setState({ participantsError: 'Failed to parse participants data.' });
        }
      }
      // Compute total cost if cost field exists (numeric)
      const totalCost = participantsArray.reduce((sum, p) => {
        if (p && typeof p === 'object') {
          const val = parseFloat(p.cost);
            if (!isNaN(val)) return sum + val;
        }
        return sum;
      }, 0);
      this.setState({ participants: participantsArray, participantsLoading: false, totalCost });
      console.log("Participants:", participantsArray, "Total Cost:", totalCost);
    } catch (error) {
      console.error("Error fetching participants:", error);
      this.setState(prev => ({ participantsError: error.message, participantsLoading: false, rawResponse: prev.rawResponse }));
    }
  };

  render() {
    let meetingId = this.state.meetingId ?? "";
    let meetingId_decoded = this.state.meetingIdDecoded ?? "";
    let userPrincipleName = this.state.userName ?? "";
  const { participants, participantsLoading, participantsError, totalCost } = this.state;

    return (
      <div>
        <h1>Meeting Costs App</h1>
        
        <p>This app provides an estimated total cost of this meeting based on its duration and the assumed cost of participants by job title.
Please note: these estimates are only rough approximations and may not accurately reflect the actual costs. 
        </p>
{/*         <h3>User Principal Name:</h3>
        <p>{userPrincipleName}</p>
        <h3>Meeting ID Original:</h3>
        <p>{meetingId}</p>
        <h3>Meeting ID Decoded (not base64):</h3>
        <p>{meetingId_decoded}</p> */}
        

        <h3>Participants:</h3>
        {participantsLoading && <p>Loading participants...</p>}
        {participantsError && (
          <div style={{ color: 'red' }}>
            <p>Error: {participantsError}</p>
            {this.state.rawResponse && (
              <pre style={{ whiteSpace: 'pre-wrap', maxHeight: 200, overflow: 'auto', background: '#f5f5f5', padding: 8 }}>
                {this.state.rawResponse}
              </pre>
            )}
          </div>
        )}
        {!participantsLoading && !participantsError && participants && participants.length === 0 && (
          <p>No participants found.</p>
        )}
        {!participantsLoading && !participantsError && participants && participants.length > 0 && (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ borderCollapse: 'collapse', width: '100%' }}>
              <thead>
                <tr style={{ background: '#f0f0f0' }}>
                  <th style={{ textAlign: 'left', padding: '6px', border: '1px solid #ccc' }}>Name / UPN</th>
                  <th style={{ textAlign: 'left', padding: '6px', border: '1px solid #ccc' }}>Role</th>
                  <th style={{ textAlign: 'left', padding: '6px', border: '1px solid #ccc' }}>Job Title</th>
                  <th style={{ textAlign: 'right', padding: '6px', border: '1px solid #ccc' }}>Cost</th>
                </tr>
              </thead>
              <tbody>
                {participants.map((p, idx) => {
                  if (typeof p === 'string') {
                    return (
                      <tr key={idx}>
                        <td style={{ padding: '6px', border: '1px solid #ddd' }}>{p}</td>
                        <td style={{ padding: '6px', border: '1px solid #ddd' }}>-</td>
                        <td style={{ padding: '6px', border: '1px solid #ddd' }}>-</td>
                        <td style={{ padding: '6px', border: '1px solid #ddd', textAlign: 'right' }}>-</td>
                      </tr>
                    );
                  }
                  const name = p.name || p.displayName || p.upn || p.userPrincipalName || p.email || '';
                  const role = p.role || p.userRole || '';
                  const jobTitle = p.jobTitle || p.jobtitle || p.title || '';
                  const costVal = p.cost !== undefined ? parseFloat(p.cost) : undefined;
                  const costDisplay = costVal !== undefined && !isNaN(costVal) ? costVal.toFixed(2) : '';
                  return (
                    <tr key={idx}>
                      <td style={{ padding: '6px', border: '1px solid #ddd' }}>{name}</td>
                      <td style={{ padding: '6px', border: '1px solid #ddd' }}>{role}</td>
                      <td style={{ padding: '6px', border: '1px solid #ddd' }}>{jobTitle}</td>
                      <td style={{ padding: '6px', border: '1px solid #ddd', textAlign: 'right' }}>{costDisplay}</td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr style={{ background: '#fafafa', fontWeight: 'bold' }}>
                  <td style={{ padding: '6px', border: '1px solid #ccc' }} colSpan={3}>Total Cost</td>
                  <td style={{ padding: '6px', border: '1px solid #ccc', textAlign: 'right' }}>{totalCost.toFixed(2)}</td>
                </tr>
              </tfoot>
            </table>
          </div>
        )}

        {/*this.state.rawResponse && !participantsLoading && (
          <div style={{ marginTop: 20 }}>
            <h4>Raw Response</h4>
            <pre style={{ whiteSpace: 'pre-wrap', maxHeight: 300, overflow: 'auto', background: '#eef', padding: 8 }}>
              {this.state.rawResponse}
            </pre>
          </div>
        )*/}

        

        <MediaQuery maxWidth={280}>
          <h3>This is the side panel</h3>
          <a href="https://docs.microsoft.com/en-us/microsoftteams/platform/apps-in-teams-meetings/teams-apps-in-meetings">
            Need more info, open this document in new tab or window.
          </a>
        </MediaQuery>
      </div>
    );
  }
}

export default Tab;