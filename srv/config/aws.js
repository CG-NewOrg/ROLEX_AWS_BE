const { S3Client } = require('@aws-sdk/client-s3');
const { getDestination } = require('@sap-cloud-sdk/connectivity');

async function getObjectStoreConfig(options = {}) {
  const { destinationName = 'object-store-dest' } = options;

  // Fetch destination
  const destination = await getDestination(
    { destinationName },
    { useCache: false }
  );
  
  if (!destination) {
    throw new Error(`Destination '${destinationName}' not found`);
  }

  // Get additional properties from destination
  const props = destination.originalProperties?.destinationConfiguration || 
                destination.originalProperties || 
                {};
  
  // BasicAuthentication: credentials from username/password
  const accessKeyId = destination.username || props.User;
  const secretAccessKey = destination.password || props.Password;
  const region = props.region || props.Region;
  const bucketName = props.bucketName || props.BucketName;
  

  // Validate required fields
  const missingFields = [];
  if (!accessKeyId) missingFields.push('User (AWS Access Key ID)');
  if (!secretAccessKey) missingFields.push('Password (AWS Secret Access Key)');
  if (!region) missingFields.push('region');
  if (!bucketName) missingFields.push('bucketName');

  if (missingFields.length > 0) {
    throw new Error(`Destination '${destinationName}' missing required properties: ${missingFields.join(', ')}`);
  }

  // Create S3 client
  const s3 = new S3Client({
    region,
    credentials: {
      accessKeyId,
      secretAccessKey
    }
  });

  return {
    s3,
    bucketName,
    region,
    destinationUrl: destination.url
  };
}

module.exports = {
  getObjectStoreConfig
};