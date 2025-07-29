export class helper {
  static normailzePath = (path: string): string => {
    return path?.replace(/^\/+|\/+$/g, '') || '';
  }
}