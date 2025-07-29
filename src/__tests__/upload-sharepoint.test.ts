
import axios from 'axios';
import { uploadToSharePoint } from '../../src/core/microsoft';
import { helper } from '../../src/utils/helper';

jest.mock('axios');
jest.mock('../../src/utils/helper', () => ({
  helper: {
    normailzePath: jest.fn()
  }
}));

const mockedAxios = axios as jest.Mocked<typeof axios>;
const mockedHelper = helper as jest.Mocked<typeof helper>;

describe('uploadToSharePoint', () => {
  const dummyBuffer = Buffer.from('dummy content');
  const dummyToken = 'dummy-access-token';

  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('uploads file to root folder if folderPath is empty', async () => {
    mockedHelper.normailzePath.mockReturnValueOnce('');

    mockedAxios.put.mockResolvedValueOnce({
      status: 201,
      data: { id: '123', name: 'file.txt' }
    });

    const res = await uploadToSharePoint(
      'tenant',
      'site',
      'library-id',
      'file.txt',
      dummyBuffer,
      dummyToken,
      false,
      '' // folderPath
    );

    expect(mockedHelper.normailzePath).toHaveBeenCalledWith('');
    expect(mockedAxios.put).toHaveBeenCalledWith(
      'https://graph.microsoft.com/v1.0/drives/library-id/root:/file.txt:/content',
      dummyBuffer,
      {
        headers: {
          Authorization: `Bearer ${dummyToken}`,
          'Content-Type': 'application/octet-stream'
        }
      }
    );

    expect(res).toEqual({ id: '123', name: 'file.txt' });
  });

  it('uploads file to nested folder path', async () => {
    mockedHelper.normailzePath.mockReturnValueOnce('folder/subfolder');

    mockedAxios.put.mockResolvedValueOnce({
      status: 201,
      data: { id: 'xyz', name: 'image.png' }
    });

    const res = await uploadToSharePoint(
      'tenant',
      'site',
      'library-id',
      'image.png',
      dummyBuffer,
      dummyToken,
      false,
      '/folder/subfolder/'
    );

    expect(mockedAxios.put).toHaveBeenCalledWith(
      'https://graph.microsoft.com/v1.0/drives/library-id/root:/folder/subfolder/image.png:/content',
      dummyBuffer,
      expect.any(Object)
    );

    expect(res.name).toBe('image.png');
  });

  it('logs message if isShowLog is true', async () => {
    const consoleSpy = jest.spyOn(console, 'log').mockImplementation();

    mockedHelper.normailzePath.mockReturnValueOnce('');
    mockedAxios.put.mockResolvedValueOnce({
      status: 201,
      data: { id: 'file-id', name: 'log.txt' }
    });

    await uploadToSharePoint(
      'tenant',
      'site',
      'library-id',
      'log.txt',
      dummyBuffer,
      dummyToken,
      true // isShowLog
    );

    expect(consoleSpy).toHaveBeenCalledWith(
      `[kas-on-cloud]: File "log.txt" uploaded successfully to SharePoint`
    );

    consoleSpy.mockRestore();
  });

  it('throws error when upload fails', async () => {
    mockedHelper.normailzePath.mockReturnValueOnce('');
    mockedAxios.put.mockResolvedValueOnce({
      status: 500,
      statusText: 'Internal Server Error'
    });

    await expect(
      uploadToSharePoint(
        'tenant',
        'site',
        'library-id',
        'error.txt',
        dummyBuffer,
        dummyToken
      )
    ).rejects.toThrow('[kas-on-cloud]: Failed to upload file: Internal Server Error');
  });
});
