using RabbitMQ.Client;
using Shared;
using System.Text;
using System.Text.Json;

namespace RabbitMQ.ExcelCreate.Services
{
    public class RabbitMQPublisher
    {
        private readonly RabbitMQClientService _rabbitmqClientService;

        public RabbitMQPublisher(RabbitMQClientService rabbitmqClientService)
        {
            _rabbitmqClientService = rabbitmqClientService;
        }

        public void Publish(CreateExcelMessage createExcelMessage)
        {
            var channel = _rabbitmqClientService.Connect();

            var messageString = JsonSerializer.Serialize(createExcelMessage);
            var messageBody = Encoding.UTF8.GetBytes(messageString);

            var properties = channel.CreateBasicProperties();
            properties.Persistent = true;

            channel.BasicPublish(RabbitMQClientService.ExchangeName, RabbitMQClientService.RoutingKey, properties, messageBody);
        }
    }
}
