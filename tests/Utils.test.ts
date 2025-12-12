import { gasRequire } from 'tgas-local'
import { vi, beforeEach, expect, describe, it } from 'vitest'
import { mockSpreadsheetApp, mockUrlFetchApp, mockLogger, mockSpreadsheet, mockRange, mockUtilities, mockHtmlService, mockCacheService} from './mocks';

const mocks = {
  SpreadsheetApp: mockSpreadsheetApp,
  UrlFetchApp: mockUrlFetchApp,
  Logger: mockLogger,
  Utilities: mockUtilities,
  HtmlService: mockHtmlService,
  CacheService: mockCacheService
}

const glib = gasRequire('./src', mocks)


describe("Utils Tests", () => {
  beforeEach(() => {
    vi.resetAllMocks()
  })
  describe("GetSpreadSheetData", () => {
    it('throws and error if spreadsheetName could not be found', () => {
      mockSpreadsheet.getSheetByName.mockImplementation(() => null)
      expect(() => glib.getSpreadSheetData("Test")).toThrow(/^Could not find spreadsheet: "Test"$/)
    })
    it('returns properly formatted data for JCIDS', () => {
      const mockData = [
        ['Description', 'Code'],
        ['Desc1', 'Code1'],
        ['Desc2', 'Code2'],
        ['Desc3', 'Code3'],
        ['Desc4', 'Code4']
      ]
      const expectedData = [
        {Description: 'Desc1', Code: 'Code1'},
        {Description: 'Desc2', Code: 'Code2'},
        {Description: 'Desc3', Code: 'Code3'},
        {Description: 'Desc4', Code: 'Code4'},
      ]
      mockRange.getValues.mockReturnValue(mockData)
      const returnData = glib.getSpreadSheetData('Test')
      expect(returnData).toEqual(expectedData)
    })
    it('returns properly formatted data for Customers (with empty columns)', () => {
      const mockData = [
        ['Name', 'Address1', 'Address2', 'City', 'State', 'Zip', 'Category'],
        ['Cust1', 'Cust1Address1', '', 'Cust1City', 'Cust1State', '', ''],
        ['Cust2', '','', 'Cust2City', 'Cust2State', '', 'Cust2Category'],
        ['Cust3', 'Cust3Address1', 'Cust3Address2', 'Cust3City', 'Cust3State', 'Cust3Zip', 'Cust3Category']
      ]
      const expectedData = [
        { Name: 'Cust1', Address1: 'Cust1Address1', Address2: '', City: 'Cust1City', State: 'Cust1State', Zip: '', Category: ''},
        { Name: 'Cust2', Address1: '', Address2: '', City: 'Cust2City', State: 'Cust2State', Zip: '', Category: 'Cust2Category'},
        { Name: 'Cust3', Address1: 'Cust3Address1', Address2: 'Cust3Address2', City: 'Cust3City', State: 'Cust3State', Zip: 'Cust3Zip', Category: 'Cust3Category'}
      ]
      mockRange.getValues.mockReturnValue(mockData)
      const returnData = glib.getSpreadSheetData('Test')
      expect(returnData).toEqual(expectedData)
    })
    it('trimp whitespace for strings', () => {
      const mockData = [
        ['Description', 'Code'],
        ['Desc1', 'Code1      '],
        ['       Desc2', 1234]
      ]
      const expectedData = [
        {Description: "Desc1", Code: 'Code1'},
        {Description: 'Desc2', Code: 1234}
      ]
      mockRange.getValues.mockReturnValue(mockData)
      const returnData = glib.getSpreadSheetData("Test")
      expect(returnData).toEqual(expectedData)
    })
  })
describe('batchFetch()', () => {
    it('should call UrlFetchApp multiple times for large number of batchOptions', () => {
      let largeArray: string[] = new Array(500)
      largeArray = largeArray.fill('option')
      mockUrlFetchApp.fetchAll.mockImplementation((options: string[]) => new Array(options.length).fill({getResponseCode: () => 201, getContentText: () => "Message"}))
      const results = glib.batchFetch(largeArray)
      expect(results.length).toBe(500)
      expect(mockUrlFetchApp.fetchAll).toHaveBeenCalledTimes(Math.ceil(largeArray.length / glib.DEFAULT_BATCH_SIZE));
      expect(mockUtilities.sleep).toHaveBeenCalledTimes(Math.ceil(largeArray.length / glib.DEFAULT_BATCH_SIZE));
    })
    it('should not call sleep more than once if only making one api call', () => {
      const array = ['option', 'option', 'option']
      mockUrlFetchApp.fetchAll.mockImplementation((options: string[]) => new Array(options.length).fill({getResponseCode: () => 201, getContentText: () => "Message"}))
      const results = glib.batchFetch(array)
      expect(results.length).toBe(3)
      expect(mockUtilities.sleep).toHaveBeenCalledOnce();
    })
    it('should not make any calls if batch length is 0', () => {
      const array: string[] = []
      const results = glib.batchFetch(array)
      expect(results.length).toBe(0)
      expect(mockUrlFetchApp.fetchAll).not.toHaveBeenCalled()
    })
    it('should retry when UrlFetchApp returns with 500 timeout repsonse', () => {
      const array = ['1', '2', '3', '4', '5'];
      mockUrlFetchApp.fetchAll.mockImplementation((values: string[]) => values.map(each => ({getResponseCode: () => 201, getContentText: () => each})))
      mockUrlFetchApp.fetchAll.mockImplementationOnce((values: string[]) => {
        return values.map((each, index) => {
          return index === 0 ? {
            getResponseCode: () =>  500,
            getContentText: () => "Connection Timeout Expired."
          } :
          {
            getResponseCode: () => 201,
            getContentText: () => each
          }
        })
      })
      
      const results = glib.batchFetch(array);
      expect(results.length).toBe(array.length)
      expect(results[0].getContentText()).toBe(array[0]);
      expect(mockUtilities.sleep).toHaveBeenCalledTimes(2)
    })
    it('should retry max of 5 times when 500 response continues', () => {
      mockUrlFetchApp.fetchAll.mockImplementation((values: []) => { 
          return values.map((each, index) => {
          return index === 0 ? {
            getResponseCode: () =>  500,
            getContentText: () => "Connection Timeout Expired."
          } :
          {
            getResponseCode: () => 201,
            getContentText: () => each
          }
        })
      })
      const array = ['1', '2', '3', '4', '5'];
      const results = glib.batchFetch(array);
      expect(results.length).toBe(array.length)
      expect(results[0].getContentText()).toBe("Connection Timeout Expired.")
      expect(mockUtilities.sleep).toHaveBeenCalledTimes(6)
      expect(mockLogger.log).toHaveBeenCalledTimes(5)
    })
  })
})