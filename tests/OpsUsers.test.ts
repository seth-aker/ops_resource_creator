import { beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import { gasRequire } from "tgas-local";
import { mockAuthenticate, mockCacheService, mockHtmlService, mockLogger, mockSpreadsheetApp, mockUi, mockUrlFetchApp, mockUtilities } from "./mocks";

const mocks = {
  Logger: mockLogger,
  SpreadsheetApp: mockSpreadsheetApp,
  UrlFetchApp: mockUrlFetchApp,
  Utilities: mockUtilities,
  HtmlService: mockHtmlService,
  CacheService: mockCacheService
}

const glib = gasRequire('./src', mocks);

describe('OpsUsers tests', () => {
  const mockGetSpreadsheetData = vi.fn()
  beforeAll(() => {
    glib.getSpreadSheetData = mockGetSpreadsheetData;
    glib.authenticate = mockAuthenticate
    glib.createHeaders = vi.fn().mockImplementation(() => {});
    glib.highlightRows = vi.fn()
  })
  beforeEach(() => {
    vi.resetAllMocks()
  })

  it('exits early if getSpreadSheetData() returns with no data', () => {
    mockGetSpreadsheetData.mockReturnValue([])
    glib.CreateUsers()
    expect(glib.getSpreadSheetData).toHaveBeenCalledOnce();
    expect(mockAuthenticate).toHaveBeenCalledOnce();
    expect(mocks.Logger.log).toHaveBeenCalledWith("No data to send!")
    expect(mockUi.alert).toHaveBeenCalledWith("No data to send!");
    expect(mockUrlFetchApp.fetchAll).not.toHaveBeenCalled()
  })
  it('alerts users of successfull run when all responseCodes are 201', () => {
    mockGetSpreadsheetData.mockReturnValue([1,2,3]);
    glib.batchFetch = vi.fn().mockImplementationOnce(() => [
      {getResponseCode: () => 201},
      {getResponseCode: () => 201},
      {getResponseCode: () => 201},
    ])

    glib.CreateUsers()
    expect(mockLogger.log).nthCalledWith(1, "Row 2: Successfully created.")
    expect(mockLogger.log).nthCalledWith(2, "Row 3: Successfully created.")
    expect(mockLogger.log).nthCalledWith(3, "Row 4: Successfully created.")
    expect(mockUi.alert("All users created successfully!"));
  })
  it('returns existing rows when response codes are either 200 or 409', () => {
    glib.batchFetch = vi.fn().mockImplementationOnce(() => [
      {getResponseCode: () => 409},
      {getResponseCode: () => 200},
      {getResponseCode: () => 201},
    ])
    mockGetSpreadsheetData.mockReturnValue([1,2,3]);

    glib.CreateUsers()

    expect(mockLogger.log).nthCalledWith(1, "Row 2: Already exists in the database.")
    expect(mockLogger.log).nthCalledWith(2, "Row 3: Already exists in the database.")
    expect(mockLogger.log).nthCalledWith(3, "Row 4: Successfully created.")
    expect(glib.highlightRows).toHaveBeenCalledWith([2,3], 'yellow')
    expect(mockUi.alert).toHaveBeenCalledWith("2 users already existed in the database. Rows: 2, 3")
  })
  it('returns failed rows when response codes are <= 400 and not 409', () => {
     glib.batchFetch = vi.fn().mockImplementationOnce(() => [
      {getResponseCode: () => 400, getContentText: () => "Error 400"},
      {getResponseCode: () => 404, getContentText: () => "Error 404"},
      {getResponseCode: () => 500, getContentText: () => "Error 500"},
    ])
    mockGetSpreadsheetData.mockReturnValue([1,2,3]);

    glib.CreateUsers()

    expect(mockLogger.log).nthCalledWith(1, "User at row 2 failed with status code: 400. Error Message: Error 400")
    expect(mockLogger.log).nthCalledWith(2, "User at row 3 failed with status code: 404. Error Message: Error 404")
    expect(mockLogger.log).nthCalledWith(3, "User at row 4 failed with status code: 500. Error Message: Error 500")
    expect(glib.highlightRows).toHaveBeenCalledWith([2,3,4], 'red')
    expect(mockUi.alert).toHaveBeenCalledWith('3 users failed to be created at rows: 2, 3, 4')
  })
})