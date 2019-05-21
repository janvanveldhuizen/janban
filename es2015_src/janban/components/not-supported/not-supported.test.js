import app from '.';

describe('notSupported', () => {

  describe('NotSupportedCtrl', () => {
    let ctrl;

    beforeEach(() => {
      angular.mock.module(app);

      angular.mock.inject(($controller) => {
        ctrl = $controller('NotSupportedCtrl', {});
      });
    });

    it('should contain the sorry message', () => {
      expect(ctrl.message).toBe('Sorry, this app can only be run as the home page of a folder in Outlook');
    });
  });
});