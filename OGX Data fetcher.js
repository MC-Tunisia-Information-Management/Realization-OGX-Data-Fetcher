function fetchRealizations_oGX() {
  var graphQLQuery = `
      query AllOpportunityApplication {
        allOpportunityApplication(
            filters: {
                date_realized: { from: "2024-02-01T00:00:00Z", to: "2024-12-15T00:00:00Z" }
                person_committee: lc_id
                programmes: [7, 8, 9]
                
            }
            pagination: { per_page: 3000 }
        ) {
            data {
                person {
                    id
                    full_name
                    home_mc {
                        name
                    }
                    home_lc {
                        name
                    }
                }
                date_realized
                opportunity {
                    programme {
                        short_name_display
                    }
                    title
                    opportunity_duration_type {
                        duration_type
                    }
                    sdg_info {
                        sdg_target {
                            goal_index
                        }
                    }
                    sub_product {
                        name
                    }
                    home_mc {
                      name
                  }
                    home_lc {
                      name
                  }
                }
            }
        }
      }
    `;

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ query: graphQLQuery }),
  };

  var url = "https://gis-api.aiesec.org/graphql?access_token={access_token}";
  var response = UrlFetchApp.fetch(url, options);

  // Parse response
  var responseData = JSON.parse(response.getContentText());
  var applicationData = responseData.data.allOpportunityApplication.data;

  // Get Google Sheets spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Sheet name");

  sheet.getRange("B7:M").clearContent();

  // Write data
  var data = applicationData.map(function (application) {
    var programme = application.opportunity.programme
      ? application.opportunity.programme.short_name_display
      : "";
    var subProduct = application.opportunity.sub_product
      ? application.opportunity.sub_product.name
      : "";

    // Handle sub_product based on the programme
    if (programme === "GV") {
      subProduct = "Volunteering";
    } else if (programme === "GTe") {
      subProduct = "Teaching";
    } else {
      subProduct = subProduct ? subProduct : "N/A"; // Show original value or 'N/A' if null
    }

    return [
      application.person.id,
      application.person.full_name,
      application.person.home_mc ? application.person.home_mc.name : "N/A",
      application.person.home_lc ? application.person.home_lc.name : "N/A",
      application.opportunity.home_mc
        ? application.opportunity.home_mc.name
        : "N/A",
      application.opportunity.home_lc
        ? application.opportunity.home_lc.name
        : "N/A",
      application.date_realized,
      programme,
      application.opportunity.title,
      application.opportunity.opportunity_duration_type
        ? application.opportunity.opportunity_duration_type.duration_type
        : "N/A",
      application.opportunity.sdg_info
        ? application.opportunity.sdg_info.sdg_target.goal_index
        : "N/A",
      subProduct,
    ];
  });
  sheet.getRange("B7:M" + (data.length + 6)).setValues(data); // Adjusted range
}
