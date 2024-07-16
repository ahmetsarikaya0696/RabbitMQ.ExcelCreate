using RabbitMQ.Client;

namespace ExcelCreateWorkerService.Services
{
    public class RabbitMQClientService
    {
        private readonly ConnectionFactory _connectionFactory;
        private IConnection _connection;
        private IModel _channel;
        public static string QueueName = "queue-excel-file";
        private readonly ILogger<RabbitMQClientService> _logger;

        public RabbitMQClientService(ConnectionFactory connectionFactory, ILogger<RabbitMQClientService> logger)
        {
            _connectionFactory = connectionFactory;
            _logger = logger;
        }

        public IModel Connect()
        {
            _connection = _connectionFactory.CreateConnection();

            if (_channel is { IsOpen: true })
            {
                return _channel;
            }

            _channel = _connection.CreateModel();

            _logger.LogInformation("RabbitMQ ile bağlantı kuruldu!");

            return _channel;
        }

        public void Dispose()
        {
            _channel?.Dispose();
            _channel?.Dispose();

            _connection?.Close();
            _connection?.Dispose();

            _logger.LogInformation("RabbitMQ ile bağlantı koptu!");
        }
    }
}
