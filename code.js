(function (window, undefined) {
  window.Asc.plugin.init = function () {
    const canvas = get_canvas();
    const image = get_image();

    function get_image() {
      return document.getElementById('input-image');
    }

    function get_canvas() {
      return document.getElementById('canvas');
    }

    document.getElementById('load-input').onchange = (event) => {
      image.src = URL.createObjectURL(event.target.files[0]);
      image.onload = () => {
        drawImage();
      }
    };

    function drawImage() {
      let scale = 70 / image.width
      canvas.width = 70;
      canvas.height = image.height * scale;
      var ctx = canvas.getContext('2d');
      ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
      Asc.scope.st = { "data": ctx.getImageData(0, 0, canvas.width, canvas.height).data.toString(), "width": canvas.width, "height": canvas.height };
      filling_cells()
    }

    function filling_cells() {
      window.Asc.plugin.callCommand(function () {
        var oWorksheet = Api.GetActiveSheet();
        const height = Asc.scope.st['height']
        const width = Asc.scope.st['width']
        const data = Asc.scope.st['data'].split(',')
        let dataIndex = 0
        let dataR;
        let dataG;
        let dataB;
        for (let x = 0; x < Asc.scope.st['height']; x++) {
          for (let y = 0; y < Asc.scope.st['width']; y++) {
            const alpha = ((data[dataIndex + 3])) / 255
            dataR = alpha * data[dataIndex]
            dataG = alpha * data[dataIndex + 1]
            dataB = alpha * data[dataIndex + 2]
            oWorksheet.GetRangeByNumber(x, y).SetFillColor(Api.CreateColorFromRGB(dataR, dataG, dataB));
            dataIndex = dataIndex + 4
          }
          oWorksheet.SetColumnWidth(x, 2);
        }
      });
    }
  };
})(window, undefined);    
